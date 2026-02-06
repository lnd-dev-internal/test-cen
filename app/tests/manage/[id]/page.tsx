'use client'

import { useEffect, useState, type ClipboardEvent } from 'react'
import { useParams, useRouter, useSearchParams } from 'next/navigation'
import { createClient } from '@/utils/supabase/client'
import * as XLSX from 'xlsx'

type QuestionType = 'single' | 'multiple' | 'essay'

type AnswerOption = {
  id: string
  dbId?: string
  text: string
  isCorrect: boolean
  images: string[] // Changed from image_url
}

type Question = {
  id: string
  content: string
  type: QuestionType
  images: string[] // Changed from image_url
  options: AnswerOption[]
}

type Section = 'info' | 'questions'

/* Helpers for Date */
function toDatetimeLocal(value: string | null) {
  if (!value) return ''
  const d = new Date(value)
  if (Number.isNaN(d.getTime())) return ''
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(
    d.getHours()
  )}:${pad(d.getMinutes())}`
}

function fromDatetimeLocal(value: string) {
  if (!value) return null
  const d = new Date(value)
  return Number.isNaN(d.getTime()) ? null : d.toISOString()
}

/* Helpers for Images */
const BUCKET = 'test-assets'

async function uploadImageToStorage(supabase: any, file: File, testId: string) {
  const ext = (file.name.split('.').pop() || 'png').toLowerCase()
  const path = `${testId}/${Date.now()}-${Math.random().toString(16).slice(2)}.${ext}`

  const { error } = await supabase.storage
    .from(BUCKET)
    .upload(path, file, { upsert: true, contentType: file.type })

  if (error) throw error

  if (error) throw error

  const { data } = supabase.storage.from(BUCKET).getPublicUrl(path)
  return data.publicUrl as string
}

function parseImages(val: string | null): string[] {
  if (!val) return []
  try {
    if (val.trim().startsWith('[')) {
      const parsed = JSON.parse(val)
      if (Array.isArray(parsed)) return parsed.filter((x: any) => typeof x === 'string')
    }
  } catch (e) { }
  return [val]
}

function getPastedImageFile(e: ClipboardEvent) {
  const items = e.clipboardData?.items
  if (!items) return null
  const imgItem = Array.from(items).find(it => it.type.startsWith('image/'))
  if (!imgItem) return null
  return imgItem.getAsFile()
}

const toLetter = (i: number) => String.fromCharCode(65 + i)

export default function ManageTestPage() {
  const supabase = createClient()
  const params = useParams<{ id: string }>()
  const router = useRouter()
  const searchParams = useSearchParams()
  const testId = params?.id

  const initialTab = searchParams.get('tab') === 'questions' ? 'questions' : 'info'
  const [activeSection, setActiveSection] = useState<Section>(initialTab)

  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [testStatus, setTestStatus] = useState<'draft' | 'published'>('draft')
  const [toggling, setToggling] = useState(false)

  // Export State
  const [exportModal, setExportModal] = useState<{ id: string; title: string } | null>(null)
  const [dateRange, setDateRange] = useState({ from: '', to: '' })
  const [exporting, setExporting] = useState(false)

  // Bank state
  const [showBankModal, setShowBankModal] = useState(false)
  const [bankQuestions, setBankQuestions] = useState<any[]>([])
  const [bankCategories, setBankCategories] = useState<any[]>([])
  const [loadingBank, setLoadingBank] = useState(false)
  const [selectedBankCat, setSelectedBankCat] = useState<string | 'all'>('all')
  const [bankSearch, setBankSearch] = useState('')

  /* ===== SMART BUILD ===== */
  const [isSmartModalOpen, setIsSmartModalOpen] = useState(false)
  const [smartConfig, setSmartConfig] = useState({
    easy: 0,
    medium: 0,
    hard: 0,
    categoryId: 'all'
  })

  const openBank = async () => {
    setShowBankModal(true)
    setLoadingBank(true)
    const { data: cats } = await supabase.from('question_bank_categories').select('*').order('name')
    // Increased limit to 5000 and using join for better performance/reliability
    const { data: qs } = await supabase
      .from('question_bank')
      .select('*, question_bank_answers(*)')
      .order('created_at', { ascending: false })
      .limit(5000)
    setBankCategories(cats || [])
    setBankQuestions(qs || [])
    setLoadingBank(false)
  }

  const importQuestion = (bq: any) => {
    const newQ: Question = {
      id: `new-${Date.now()}-${Math.random()}`,
      content: bq.content,
      type: bq.type as QuestionType,
      images: bq.images || [],
      options: bq.type === 'essay' ? [] : (bq.question_bank_answers || []).map((ba: any, idx: number) => ({
        id: toLetter(idx),
        text: ba.content || '',
        isCorrect: !!ba.is_correct,
        images: ba.images || []
      }))
    }
    setQuestions(prev => [...prev, newQ])
  }

  const handleSmartBuild = async () => {
    setIsSmartModalOpen(false)
    setSaving(true)

    const { data: qs } = await supabase.from('question_bank').select('*')
    const { data: ans } = await supabase.from('question_bank_answers').select('*')

    const mappedQs = (qs || []).map((q: any) => ({
      ...q,
      answers: (ans || []).filter((a: any) => a.question_id === q.id)
    }))

    const filtered = mappedQs.filter((q: any) => smartConfig.categoryId === 'all' || q.category_id === smartConfig.categoryId)

    const pick = (diff: string, count: number) => {
      const pool = filtered.filter((q: any) => (q.difficulty || 'Easy') === diff).sort(() => Math.random() - 0.5)
      return pool.slice(0, count)
    }

    const selected = [
      ...pick('Easy', smartConfig.easy),
      ...pick('Medium', smartConfig.medium),
      ...pick('Hard', smartConfig.hard)
    ]

    const newQs: Question[] = selected.map((q: any, i: number) => ({
      id: `new-smart-${Date.now()}-${i}`,
      content: q.content,
      type: q.type as QuestionType,
      images: q.images || [],
      options: q.answers.map((a: any, ai: number) => ({
        id: toLetter(ai),
        text: a.content,
        isCorrect: a.is_correct,
        images: a.images || [],
      }))
    }))

    setQuestions(prev => [...prev, ...newQs])
    setSaving(false)
    alert(`✅ Đã bốc ngẫu nhiên ${newQs.length} câu hỏi!`)
  }

  /* ===== DATA STATE ===== */
  const [form, setForm] = useState({
    name: '',
    description: '',
    passScore: 80,
    unlimitedTime: true,
    timeMinutes: 60,
    validFrom: '',
    validTo: '',
    successMessage: '',
    failMessage: '',
    allowReview: true,
    maxViolations: 0,
    shuffleQuestions: false,
    shuffleAnswers: false,
  })

  const [questions, setQuestions] = useState<Question[]>([])

  /* ===== LOAD ===== */
  useEffect(() => {
    if (!testId) return

    const load = async () => {
      setLoading(true)

      // 1. Load Test Info
      const { data: t, error: tErr } = await supabase
        .from('tests')
        .select('*')
        .eq('id', testId)
        .single()

      if (tErr) {
        alert(tErr.message)
        setLoading(false)
        return
      }

      // Save status
      setTestStatus(t.status ?? 'draft')

      // Map Info
      const timeLimitOn =
        typeof t.time_limit === 'boolean'
          ? t.time_limit
          : Number(t.time_limit ?? 0) === 1
      const unlimited = !timeLimitOn
      const duration = Number(t.duration_minutes ?? 0)

      setForm({
        name: t.title ?? '',
        description: t.description ?? '',
        passScore: Number(t.pass_score ?? 80),
        unlimitedTime: unlimited,
        timeMinutes: unlimited ? 60 : duration || 60,
        validFrom: toDatetimeLocal(t.valid_from),
        validTo: toDatetimeLocal(t.valid_to),
        successMessage: t.success_message ?? '',
        failMessage: t.fail_message ?? '',
        allowReview: !!t.allow_review,
        maxViolations: Number(t.max_violations ?? 0),
        shuffleQuestions: !!t.shuffle_questions,
        shuffleAnswers: !!t.shuffle_answers,
      })

      // 2. Load Questions
      const { data: qs, error: qErr } = await supabase
        .from('questions')
        .select('id, content, type, image_url')
        .eq('test_id', testId)
        .order('id', { ascending: true })

      if (qErr) {
        alert(qErr.message)
      } else {
        const qIds = (qs ?? []).map((q: any) => q.id)
        const rawByQ: Record<string, any[]> = {}

        if (qIds.length) {
          const { data: ans, error: aErr } = await supabase
            .from('answers')
            .select('id, question_id, content, is_correct, image_url')
            .in('question_id', qIds)
            .order('id', { ascending: true })

          if (aErr) alert(aErr.message)
          else {
            for (const a of ans ?? []) {
              if (!rawByQ[a.question_id]) rawByQ[a.question_id] = []
              rawByQ[a.question_id].push(a)
            }
          }
        }

        const mapped: Question[] = (qs ?? []).map((q: any) => {
          const rawAnswers = rawByQ[q.id] ?? []
          const options: AnswerOption[] =
            q.type === 'essay'
              ? []
              : rawAnswers.map((a, idx) => ({
                id: toLetter(idx),
                dbId: a.id,
                text: a.content ?? '',
                isCorrect: !!a.is_correct,
                images: parseImages(a.image_url),
              }))

          // Fix empty options if non-essay
          if (q.type !== 'essay' && options.length === 0) {
            options.push(
              { id: 'A', text: '', isCorrect: false, images: [] },
              { id: 'B', text: '', isCorrect: false, images: [] }
            )
          }

          return {
            id: q.id,
            content: q.content ?? '',
            type: q.type as QuestionType,
            images: parseImages(q.image_url),
            options,
          }
        })
        setQuestions(mapped)
      }

      setLoading(false)
    }

    load()
  }, [testId])

  useEffect(() => {
    const fetchCats = async () => {
      const { data } = await supabase.from('question_bank_categories').select('*').order('name')
      setBankCategories(data || [])
    }
    fetchCats()
  }, [])

  /* ===== TOGGLE PUBLISH ===== */
  const togglePublish = async () => {
    if (!testId) return
    const nextStatus = testStatus === 'published' ? 'draft' : 'published'
    setToggling(true)
    try {
      const { error } = await supabase
        .from('tests')
        .update({ status: nextStatus })
        .eq('id', testId)
      if (error) throw error
      setTestStatus(nextStatus)
      alert(nextStatus === 'published' ? '✅ Đã xuất bản!' : '✅ Đã chuyển về nháp!')
    } catch (err: any) {
      alert(err.message || 'Lỗi')
    } finally {
      setToggling(false)
    }
  }


  /* ===== EXPORT ===== */
  const handleExport = async (type: 'csv' | 'excel' | 'google') => {
    if (!exportModal) return
    setExporting(true)

    try {
      let query = supabase
        .from('test_submissions')
        .select('*')
        .eq('test_id', exportModal.id)
        .order('created_at', { ascending: false })

      if (dateRange.from) query = query.gte('created_at', new Date(dateRange.from).toISOString())
      if (dateRange.to) query = query.lte('created_at', new Date(dateRange.to).toISOString())

      const { data, error } = await query

      if (error) throw error
      if (!data || data.length === 0) {
        alert('Không có dữ liệu trong khoảng thời gian này')
        setExporting(false)
        return
      }

      // Map data
      const rows = data.map((s: any) => ({
        'ID': s.id,
        'Họ tên': s.candidate_name,
        'Điểm số': s.score_percent,
        'Số câu đúng': s.correct_count,
        'Tổng câu': s.total_count,
        'Kết quả': s.passed ? 'ĐẠT' : 'KHÔNG ĐẠT',
        'Thời gian làm bài (giây)': s.duration_seconds,
        'Số lần vi phạm': s.violation_count,
        'Ngày nộp': new Date(s.created_at).toLocaleString('vi-VN'),
      }))

      const workSheet = XLSX.utils.json_to_sheet(rows)
      const workBook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workBook, workSheet, "Submissions")

      if (type === 'google') {
        const csv = XLSX.utils.sheet_to_csv(workSheet, { FS: '\t' })
        await navigator.clipboard.writeText(csv)
        alert('✅ Đã copy dữ liệu! Bạn có thể paste trực tiếp vào Google Sheets.')
      } else if (type === 'csv') {
        XLSX.writeFile(workBook, `Report_${exportModal.title}_${Date.now()}.csv`)
      } else {
        XLSX.writeFile(workBook, `Report_${exportModal.title}_${Date.now()}.xlsx`)
      }

    } catch (err: any) {
      alert('Lỗi xuất dữ liệu: ' + err.message)
    } finally {
      setExporting(false)
    }
  }

  /* ===== SAVE ALL ===== */
  const saveAll = async () => {
    if (!testId) return

    // Check if published
    if (testStatus === 'published') {
      return alert('⚠️ Không thể chỉnh sửa bài test đang xuất bản. Vui lòng ngừng xuất bản trước!')
    }

    // Validate Info
    if (!form.name.trim()) return alert('Chưa nhập tên bài kiểm tra')

    // Validate Questions
    if (questions.length === 0) return alert('Chưa có câu hỏi')
    for (let i = 0; i < questions.length; i++) {
      const q = questions[i]
      if (!q.content?.trim() && q.images.length === 0) {
        return alert(`Câu ${i + 1}: chưa nhập nội dung hoặc hình ảnh`)
      }
      if (q.type !== 'essay') {
        if (q.options.length < 2) return alert(`Câu ${i + 1}: cần ít nhất 2 đáp án`)
        if (q.type === 'single' && !q.options.some(o => o.isCorrect)) {
          return alert(`Câu ${i + 1}: chọn 1 đáp án đúng`)
        }
      }
    }

    setSaving(true)

    try {
      // 1. Update Info
      const infoPayload = {
        title: form.name.trim(),
        description: form.description?.trim() || null,
        pass_score: Number(form.passScore),
        time_limit: form.unlimitedTime ? 0 : 1,
        duration_minutes: form.unlimitedTime ? 0 : Number(form.timeMinutes),
        valid_from: form.validFrom ? fromDatetimeLocal(form.validFrom) : null,
        valid_to: form.validTo ? fromDatetimeLocal(form.validTo) : null,
        success_message: form.successMessage?.trim() || null,
        fail_message: form.failMessage?.trim() || null,
        allow_review: !!form.allowReview,
        max_violations: Number(form.maxViolations),
        shuffle_questions: !!form.shuffleQuestions,
        shuffle_answers: !!form.shuffleAnswers,
      }

      const { error: infoErr } = await supabase
        .from('tests')
        .update(infoPayload)
        .eq('id', testId)

      if (infoErr) throw infoErr

      // 2. Update Questions (Diffing)
      // Note: We process one by one to keep it simple, though batching is faster.
      const qsCopy = [...questions]
      const currentQIds = []

      for (let i = 0; i < qsCopy.length; i++) {
        const q = qsCopy[i]
        const isNew = q.id.startsWith('new-')

        const correctAnsStr =
          q.type === 'essay'
            ? ''
            : q.options
              .filter(o => o.isCorrect)
              .map(o => o.id) // This is just Letter, not DB ID. Correct logic might need DB ID if required? 
              // Actually, existing code uses `o.id` (A, B...) for `correct_answer` column if it's text based?
              // Wait, let's check `create/page.tsx`:
              // `correct_answer: correctAnswer` where `correctAnswer` is ID (A, B...).
              // So yes, `correct_answer` stores the Letter ID (like 'A'), NOT the UUID.
              .join(',')

        const qPayload = {
          test_id: testId,
          content: q.content ?? '',
          type: q.type,
          correct_answer: correctAnsStr,
          options: q.type === 'essay' ? [] : q.options.map(o => o.id), // A, B, C...
          image_url: q.images.length > 0 ? JSON.stringify(q.images) : null,
        }

        let realQId = q.id

        if (isNew) {
          const { data: insQ, error: insErr } = await supabase
            .from('questions')
            .insert(qPayload)
            .select('id')
            .single()
          if (insErr) throw insErr
          realQId = insQ.id
          qsCopy[i].id = realQId // Update local ID
        } else {
          const { error: updErr } = await supabase
            .from('questions')
            .update({
              content: qPayload.content,
              type: qPayload.type,
              correct_answer: qPayload.correct_answer,
              options: qPayload.options,
              image_url: qPayload.image_url,
            })
            .eq('id', realQId)
          if (updErr) throw updErr
        }
        currentQIds.push(realQId)

        // Handle Answers
        if (q.type === 'essay') {
          // Delete old answers
          await supabase.from('answers').delete().eq('question_id', realQId)
        } else {
          // Delete ALL old answers and re-insert is Easiest way to handle re-ordering and updates safely
          // BUT `answers` table has ID. If we delete, we lose history?
          // If we want to preserve IDs, we need diffing.
          // For simplicity and robustness given previous code style:
          // Existing `questions/page.tsx` DELETED all answers and RE-INSERTED.
          // See lines 250-260 in Step 35.
          await supabase.from('answers').delete().eq('question_id', realQId)

          const ansPayload = q.options.map(o => ({
            question_id: realQId,
            content: o.text ?? '',
            is_correct: !!o.isCorrect,
            image_url: o.images.length > 0 ? JSON.stringify(o.images) : null,
          }))

          if (ansPayload.length > 0) {
            const { error: ansErr } = await supabase.from('answers').insert(ansPayload)
            if (ansErr) throw ansErr
          }
        }
      }

      // Delete removed questions
      // We did not track "deleted" questions explicitly, but we have `currentQIds`.
      // Any question currently in DB for this test that is NOT in `currentQIds` should be deleted.
      // However, `saveAll` in `questions/page.tsx` didn't implement this "Delete Others".
      // It relied on `deleteQuestion` function immediately deleting from DB.
      // Since we want "Create" style where we might delete locally?
      // "Create" doesn't have DB IDs yet.
      // "Manage" usually deletes immediately.
      // If I want "Manage" to be like "Create" (Save at end), I should defer deletions?
      // But standard `deleteQuestion` in `questions/page.tsx` was immediate.
      // I will keep `deleteQuestion` IMMEDIATE for existing Qs to avoid complex state tracking of "deletedIds".
      // This is slightly hybrids, but safe.

      setQuestions(qsCopy)
      alert('✅ Đã lưu thay đổi!')
    } catch (err: any) {
      console.error(err)
      alert(err.message || 'Lưu thất bại')
    } finally {
      setSaving(false)
    }
  }

  /* ===== DELETE QUESTION ===== */
  const deleteQuestion = async (index: number) => {
    if (testStatus === 'published') {
      return alert('⚠️ Không thể xóa câu hỏi khi bài test đang xuất bản!')
    }

    const q = questions[index]
    const isNew = q.id.startsWith('new-')

    if (!confirm('Bạn có chắc muốn xoá câu hỏi này?')) return

    if (isNew) {
      setQuestions(prev => prev.filter((_, i) => i !== index))
    } else {
      // Delete immediately from DB
      try {
        const { error } = await supabase.from('questions').delete().eq('id', q.id)
        if (error) throw error
        setQuestions(prev => prev.filter((_, i) => i !== index))
      } catch (e: any) {
        alert(e.message)
      }
    }
  }

  if (loading) return <div className="p-8">Đang tải...</div>

  const isPublished = testStatus === 'published'

  return (
    <div className="min-h-screen w-full bg-white text-gray-900">
      <div className="max-w-6xl mx-auto px-8 py-10 space-y-8 pb-32">
        {/* Header with Title and Publish Toggle */}
        <div className="flex items-center justify-between">
          <h1 className="text-3xl font-bold">Quản lý bài kiểm tra</h1>
          <button
            onClick={togglePublish}
            disabled={toggling}
            className={`px-6 py-2.5 rounded-lg font-semibold transition-colors ${isPublished
              ? 'bg-orange-500 text-white hover:bg-orange-600'
              : 'bg-green-500 text-white hover:bg-green-600'
              } disabled:opacity-50`}
          >
            {toggling ? 'Đang...' : isPublished ? '🔒 Ngừng xuất bản' : '✅ Xuất bản'}
          </button>

          {/* Export Button */}
          <button
            onClick={() => {
              if (!testId) return
              const now = new Date()
              now.setMinutes(now.getMinutes() - now.getTimezoneOffset())
              const start = new Date()
              start.setDate(start.getDate() - 30)
              start.setMinutes(start.getMinutes() - start.getTimezoneOffset())

              setExportModal({ id: testId, title: form.name })
              setDateRange({
                from: start.toISOString().slice(0, 16),
                to: now.toISOString().slice(0, 16)
              })
            }}
            className="ml-3 px-6 py-2.5 rounded-lg font-semibold bg-teal-500 text-white hover:bg-teal-600 transition-colors"
          >
            Xuất dữ liệu
          </button>
        </div>

        {/* Warning Banner when Published */}
        {isPublished && (
          <div className="bg-orange-50 border-l-4 border-orange-500 p-4 rounded">
            <div className="flex items-start">
              <div className="flex-shrink-0">
                <svg className="h-5 w-5 text-orange-500" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
                </svg>
              </div>
              <div className="ml-3">
                <p className="text-sm font-semibold text-orange-800">
                  🔒 Bài test đang ở trạng thái xuất bản
                </p>
                <p className="text-sm text-orange-700 mt-1">
                  Không thể chỉnh sửa khi đang xuất bản. Nhấn nút "Ngừng xuất bản" ở trên để có thể chỉnh sửa.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* TABS */}
        <div className="flex gap-8 border-b border-gray-200">
          <button
            onClick={() => setActiveSection('info')}
            className={`pb-3 font-semibold ${activeSection === 'info'
              ? 'border-b-2 border-[#ff5200] text-[#ff5200]'
              : 'text-gray-500'
              }`}
          >
            Thông tin cơ bản
          </button>

          <button
            onClick={() => setActiveSection('questions')}
            className={`pb-3 font-semibold ${activeSection === 'questions'
              ? 'border-b-2 border-[#ff5200] text-[#ff5200]'
              : 'text-gray-500'
              }`}
          >
            Câu hỏi ({questions.length})
          </button>
        </div>

        {/* ================= SECTION 1: INFO ================= */}
        {activeSection === 'info' && (
          <div className="border border-gray-200 rounded-xl p-8 space-y-8">
            <div className="grid grid-cols-2 gap-8">
              <Field label="Tên bài kiểm tra">
                <input
                  value={form.name}
                  onChange={e => setForm({ ...form, name: e.target.value })}
                  disabled={isPublished}
                  className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>

              <Field label="Điểm đạt (%)">
                <input
                  type="number"
                  value={form.passScore}
                  onChange={e =>
                    setForm({ ...form, passScore: Number(e.target.value) })
                  }
                  disabled={isPublished}
                  className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>
            </div>

            <Field label="Số lần vi phạm tối đa (0 = không giới hạn)">
              <input
                type="number"
                min="0"
                value={form.maxViolations}
                onChange={e =>
                  setForm({ ...form, maxViolations: Number(e.target.value) })
                }
                disabled={isPublished}
                className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
              />
              <p className="text-sm text-gray-500 mt-1">
                Vi phạm bao gồm: chuyển tab, chụp màn hình, thu nhỏ màn hình. Khi vượt quá số lần cho phép, bài làm sẽ bị khóa.
              </p>
            </Field>

            <Field label="Mô tả">
              <textarea
                value={form.description}
                onChange={e =>
                  setForm({ ...form, description: e.target.value })
                }
                disabled={isPublished}
                className="w-full min-h-[110px] px-3 py-2 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
              />
            </Field>

            <div className="grid grid-cols-2 gap-8">
              <div className="space-y-3">
                <label className="flex items-center gap-2 text-sm">
                  <input
                    type="checkbox"
                    checked={form.unlimitedTime}
                    onChange={e =>
                      setForm({ ...form, unlimitedTime: e.target.checked })
                    }
                    disabled={isPublished}
                  />
                  Không giới hạn thời gian
                </label>

                {!form.unlimitedTime && (
                  <input
                    type="number"
                    placeholder="Thời gian (phút)"
                    value={form.timeMinutes}
                    onChange={e =>
                      setForm({
                        ...form,
                        timeMinutes: Number(e.target.value),
                      })
                    }
                    disabled={isPublished}
                    className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                  />
                )}
              </div>

              <label className="flex items-center gap-2 text-sm">
                <input
                  type="checkbox"
                  checked={form.allowReview}
                  onChange={e =>
                    setForm({ ...form, allowReview: e.target.checked })
                  }
                  disabled={isPublished}
                />
                Cho phép xem lại bài làm
              </label>

              <div className="flex flex-col gap-3">
                <label className="flex items-center gap-2 text-sm font-semibold">
                  <input
                    type="checkbox"
                    checked={form.shuffleQuestions}
                    onChange={e =>
                      setForm({ ...form, shuffleQuestions: e.target.checked })
                    }
                    disabled={isPublished}
                  />
                  Đảo thứ tự câu hỏi
                </label>
                <label className="flex items-center gap-2 text-sm font-semibold">
                  <input
                    type="checkbox"
                    checked={form.shuffleAnswers}
                    onChange={e =>
                      setForm({ ...form, shuffleAnswers: e.target.checked })
                    }
                    disabled={isPublished}
                  />
                  Đảo thứ tự câu trả lời
                </label>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-8">
              <Field label="Hiệu lực từ">
                <input
                  type="datetime-local"
                  value={form.validFrom}
                  onChange={e => setForm({ ...form, validFrom: e.target.value })}
                  disabled={isPublished}
                  className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>

              <Field label="Đến">
                <input
                  type="datetime-local"
                  value={form.validTo}
                  onChange={e => setForm({ ...form, validTo: e.target.value })}
                  disabled={isPublished}
                  className="w-full h-11 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>
            </div>

            <div className="grid grid-cols-2 gap-8">
              <Field label="Thông báo khi đạt">
                <textarea
                  value={form.successMessage}
                  onChange={e =>
                    setForm({ ...form, successMessage: e.target.value })
                  }
                  disabled={isPublished}
                  className="w-full min-h-[90px] px-3 py-2 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>

              <Field label="Thông báo khi chưa đạt">
                <textarea
                  value={form.failMessage}
                  onChange={e =>
                    setForm({ ...form, failMessage: e.target.value })
                  }
                  disabled={isPublished}
                  className="w-full min-h-[90px] px-3 py-2 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                />
              </Field>
            </div>
          </div>
        )}

        {/* ================= SECTION 2: QUESTIONS ================= */}
        {activeSection === 'questions' && (
          <div className="border border-gray-200 rounded-xl p-8 space-y-8">


            {questions.map((q, qi) => (
              <div key={q.id} className="border border-gray-200 rounded-xl">
                <div className="flex justify-between items-center px-6 py-4 bg-gray-50 border-b">
                  <div className="font-semibold">Câu {qi + 1}</div>

                  <div className="flex items-center gap-3">
                    <select
                      value={q.type}
                      onChange={e => {
                        const copy = [...questions]
                        copy[qi].type = e.target.value as QuestionType
                        if (copy[qi].type === 'essay') copy[qi].options = []
                        if (copy[qi].type !== 'essay' && copy[qi].options.length === 0) {
                          copy[qi].options = [
                            { id: 'A', text: '', isCorrect: false, images: [] },
                            { id: 'B', text: '', isCorrect: false, images: [] },
                          ]
                        }
                        setQuestions(copy)
                      }}
                      disabled={isPublished}
                      className="h-10 px-3 border border-gray-300 rounded-lg bg-white text-gray-900"
                    >
                      <option value="single">1 đáp án</option>
                      <option value="multiple">Nhiều đáp án</option>
                      <option value="essay">Tự luận</option>
                    </select>

                    <button
                      onClick={() => deleteQuestion(qi)}
                      disabled={isPublished}
                      className="text-red-500 hover:text-red-700 font-medium ml-3 disabled:text-gray-400 disabled:cursor-not-allowed"
                      title="Xóa câu hỏi này"
                    >
                      Xóa
                    </button>
                  </div>
                </div>

                <div className="p-6 space-y-6">
                  {/* QUESTION CONTENT + IMAGE PASTE */}
                  <div className="space-y-2">
                    <textarea
                      placeholder="Nội dung câu hỏi (Paste ảnh vào đây để upload)"
                      value={q.content}
                      onChange={e => {
                        const copy = [...questions]
                        copy[qi].content = e.target.value
                        setQuestions(copy)
                      }}
                      onPaste={async e => {
                        if (!testId) return
                        if (isPublished) {
                          e.preventDefault()
                          return alert('⚠️ Không thể upload ảnh khi bài test đang xuất bản!')
                        }
                        const file = getPastedImageFile(e)
                        if (!file) return
                        e.preventDefault()
                        try {
                          const url = await uploadImageToStorage(supabase, file, testId)
                          const copy = [...questions]
                          copy[qi].images = [...(copy[qi].images || []), url]
                          setQuestions(copy)
                        } catch (err: any) {
                          alert(err?.message ?? 'Upload ảnh thất bại')
                        }
                      }}
                      disabled={isPublished}
                      className="w-full min-h-[120px] px-3 py-2 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                    />
                    {q.images.length > 0 && (
                      <div className="flex flex-wrap gap-2">
                        {q.images.map((img, imgIdx) => (
                          <div key={imgIdx} className="relative inline-block">
                            <img src={img} alt="Question" className="max-h-64 rounded border" />
                            <button
                              onClick={() => {
                                const copy = [...questions]
                                copy[qi].images = copy[qi].images.filter((_, i) => i !== imgIdx)
                                setQuestions(copy)
                              }}
                              disabled={isPublished}
                              className="absolute top-1 right-1 bg-red-500 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs disabled:bg-gray-400 disabled:cursor-not-allowed hover:bg-red-600"
                              title="Xoá ảnh"
                            >
                              ✕
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {q.type !== 'essay' && (
                    <div className="space-y-4">
                      {q.options.map((o, oi) => (
                        <div key={o.id} className="flex items-center gap-3">
                          <input
                            type={q.type === 'single' ? 'radio' : 'checkbox'}
                            checked={o.isCorrect}
                            onChange={() => {
                              const copy = [...questions]
                              if (q.type === 'single') {
                                copy[qi].options.forEach(x => (x.isCorrect = false))
                              }
                              copy[qi].options[oi].isCorrect = !o.isCorrect
                              setQuestions(copy)
                            }}
                            disabled={isPublished}
                          />

                          <div className="flex-1 space-y-2">
                            <input
                              placeholder={`Đáp án ${o.id} (Paste ảnh vào đây)`}
                              value={o.text}
                              onChange={e => {
                                const copy = [...questions]
                                copy[qi].options[oi].text = e.target.value
                                setQuestions(copy)
                              }}
                              onPaste={async e => {
                                if (!testId) return
                                if (isPublished) {
                                  e.preventDefault()
                                  return alert('⚠️ Không thể upload ảnh khi bài test đang xuất bản!')
                                }
                                const file = getPastedImageFile(e)
                                if (!file) return
                                e.preventDefault()
                                try {
                                  const url = await uploadImageToStorage(supabase, file, testId)
                                  const copy = [...questions]
                                  const opts = copy[qi].options
                                  opts[oi].images = [...(opts[oi].images || []), url]
                                  setQuestions(copy)
                                } catch (err: any) {
                                  alert(err?.message ?? 'Upload ảnh thất bại')
                                }
                              }}
                              disabled={isPublished}
                              className="w-full h-10 px-3 border border-gray-300 rounded-lg bg-white text-gray-900 disabled:bg-gray-100 disabled:cursor-not-allowed"
                            />
                            {o.images && o.images.length > 0 && (
                              <div className="flex flex-wrap gap-2 mt-2">
                                {o.images.map((img, imgIdx) => (
                                  <div key={imgIdx} className="relative inline-block">
                                    <img src={img} alt="Answer" className="max-h-32 rounded border" />
                                    <button
                                      onClick={() => {
                                        const copy = [...questions]
                                        copy[qi].options[oi].images = copy[qi].options[oi].images.filter((_, i) => i !== imgIdx)
                                        setQuestions(copy)
                                      }}
                                      disabled={isPublished}
                                      className="absolute top-1 right-1 bg-red-500 text-white w-5 h-5 rounded-full flex items-center justify-center text-xs disabled:bg-gray-400 disabled:cursor-not-allowed hover:bg-red-600"
                                      title="Xoá ảnh"
                                    >
                                      ✕
                                    </button>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>

                          {q.options.length > 1 && (
                            <button
                              type="button"
                              onClick={() => {
                                const copy = [...questions]
                                copy[qi].options.splice(oi, 1)
                                // Re-index IDs
                                copy[qi].options.forEach((opt, idx) => {
                                  opt.id = String.fromCharCode(65 + idx)
                                })
                                setQuestions(copy)
                              }}
                              disabled={isPublished}
                              className="text-red-500 text-sm font-medium hover:text-red-700 hover:underline disabled:text-gray-400 disabled:cursor-not-allowed"
                            >
                              Xóa
                            </button>
                          )}
                        </div>
                      ))}

                      <button
                        onClick={() => {
                          const copy = [...questions]
                          const nextChar = String.fromCharCode(65 + copy[qi].options.length)
                          copy[qi].options.push({
                            id: nextChar,
                            text: '',
                            isCorrect: false,
                            images: []
                          })
                          setQuestions(copy)
                        }}
                        disabled={isPublished}
                        className="text-sm font-medium text-[#ff5200] disabled:text-gray-400 disabled:cursor-not-allowed"
                      >
                        + Thêm đáp án
                      </button>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ===== FIXED SAVE BAR ===== */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 px-8 py-4 flex justify-end gap-4">
        {activeSection === 'questions' && (
          <>
            <button
              onClick={openBank}
              disabled={loading || saving || isPublished}
              className="px-6 py-3 rounded-xl bg-slate-800 text-white font-bold text-lg disabled:opacity-50 active:scale-95 transition-transform flex items-center gap-2"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                <path d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4z" />
              </svg>
              Lấy từ ngân hàng
            </button>
            <button
              onClick={() => setIsSmartModalOpen(true)}
              disabled={loading || saving || isPublished}
              className="px-6 py-3 rounded-xl bg-green-600 text-white font-bold text-lg disabled:opacity-50 active:scale-95 transition-transform flex items-center gap-2"
            >
              <span>🧠</span> Smart Build
            </button>
            <button
              onClick={() =>
                setQuestions(prev => [
                  ...prev,
                  {
                    id: `new-${Date.now()}`,
                    content: '',
                    type: 'single',
                    images: [],
                    options: [
                      { id: 'A', text: '', isCorrect: false, images: [] },
                      { id: 'B', text: '', isCorrect: false, images: [] },
                    ],
                  },
                ])
              }
              disabled={loading || saving || isPublished}
              className="px-6 py-3 rounded-xl bg-[#00a0fa] text-white font-bold text-lg disabled:opacity-50 active:scale-95 transition-transform"
            >
              + Thêm câu hỏi
            </button>
          </>
        )}
        <button
          onClick={saveAll}
          disabled={saving || loading || isPublished}
          className="px-8 py-3 rounded-xl bg-[#ff5200] text-white font-bold text-lg disabled:opacity-50 active:scale-95 transition-transform"
        >
          {saving ? 'Đang lưu...' : 'Lưu tất cả'}
        </button>
      </div>

      {/* ✅ BANK MODAL */}
      {showBankModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-3xl w-full max-w-5xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
              <h2 className="text-xl font-bold text-slate-800">Ngân hàng câu hỏi</h2>
              <button
                onClick={() => setShowBankModal(false)}
                className="text-slate-400 hover:text-slate-600 transition-colors p-1.5 hover:bg-slate-200 rounded-full"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="p-6 border-b border-slate-100 bg-white flex gap-4">
              <div className="flex-1 relative">
                <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">🔍</span>
                <input
                  type="text"
                  placeholder="Tìm kiếm nội dung câu hỏi..."
                  value={bankSearch}
                  onChange={(e) => setBankSearch(e.target.value)}
                  className="w-full h-12 pl-12 pr-4 bg-slate-50 border-2 border-slate-100 rounded-xl font-bold focus:border-blue-500 outline-none transition-all"
                />
              </div>
            </div>

            <div className="flex-1 overflow-hidden flex">
              {/* Sidebar Cats */}
              <div className="w-64 border-r border-slate-100 p-4 space-y-1 overflow-y-auto">
                <button
                  onClick={() => setSelectedBankCat('all')}
                  className={`w-full text-left px-4 py-2 rounded-xl text-sm font-medium transition-all ${selectedBankCat === 'all' ? 'bg-blue-50 text-blue-600' : 'text-slate-600 hover:bg-slate-50'}`}
                >
                  Tất cả câu hỏi
                </button>
                {bankCategories.map(cat => (
                  <button
                    key={cat.id}
                    onClick={() => setSelectedBankCat(cat.id)}
                    className={`w-full text-left px-4 py-2 rounded-xl text-sm font-medium transition-all ${selectedBankCat === cat.id ? 'bg-blue-50 text-blue-600' : 'text-slate-600 hover:bg-slate-50'}`}
                  >
                    {cat.name}
                  </button>
                ))}
              </div>

              {/* Bank Questions */}
              <div className="flex-1 overflow-y-auto p-6 space-y-4">
                {loadingBank ? (
                  <div className="text-center py-12 text-slate-400 font-medium">Đang tải câu hỏi...</div>
                ) : bankQuestions
                  .filter(bq => selectedBankCat === 'all' || bq.category_id === selectedBankCat)
                  .filter(bq => bq.content.toLowerCase().includes(bankSearch.toLowerCase()))
                  .length === 0 ? (
                  <div className="text-center py-12 text-slate-400 font-medium italic">Không có câu hỏi nào khớp với tìm kiếm.</div>
                ) : (
                  bankQuestions
                    .filter(bq => selectedBankCat === 'all' || bq.category_id === selectedBankCat)
                    .filter(bq => bq.content.toLowerCase().includes(bankSearch.toLowerCase()))
                    .map(bq => {
                      const alreadyIn = questions.some(q => q.content === bq.content)
                      return (
                        <div key={bq.id} className="border border-slate-200 rounded-2xl p-4 flex justify-between items-start gap-4 hover:border-blue-200 transition-colors">
                          <div className="flex-1 space-y-2">
                            <div className={`text-[10px] font-bold uppercase tracking-wider inline-block px-2 py-0.5 rounded ${bq.type === 'essay' ? 'bg-purple-100 text-purple-600' :
                              bq.type === 'multiple' ? 'bg-blue-100 text-blue-600' : 'bg-green-100 text-green-600'
                              }`}>
                              {bq.type === 'essay' ? 'Tự luận' : bq.type === 'multiple' ? 'Nhiều đáp án' : '1 đáp án'}
                            </div>
                            <div className="text-slate-800 font-medium line-clamp-2">{bq.content}</div>
                            {bq.images && bq.images.length > 0 && (
                              <div className="flex gap-1">
                                {bq.images.map((img: string, i: number) => (
                                  <img key={i} src={img} className="h-10 w-10 object-cover rounded border" />
                                ))}
                              </div>
                            )}
                          </div>
                          <button
                            onClick={() => {
                              importQuestion(bq)
                              alert('Đã thêm vào đề thi!')
                            }}
                            disabled={alreadyIn}
                            className={`px-4 py-2 rounded-xl text-sm font-bold transition-all ${alreadyIn ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-blue-600 text-white shadow-lg shadow-blue-100 hover:brightness-110 active:scale-95'
                              }`}
                          >
                            {alreadyIn ? 'Đã có' : '+ Thêm'}
                          </button>
                        </div>
                      )
                    })
                )}
              </div>
            </div>

            <div className="p-4 border-t border-slate-100 bg-slate-50/50 flex justify-end">
              <button
                onClick={() => setShowBankModal(false)}
                className="px-6 py-2 rounded-xl bg-slate-800 text-white font-bold"
              >
                Xong
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ✅ SMART BUILD MODAL */}
      {isSmartModalOpen && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md">
          <div className="bg-white rounded-[40px] w-full max-w-lg shadow-2xl p-8 space-y-8 border border-white/20">
            <div>
              <h2 className="text-3xl font-black text-slate-900 uppercase tracking-tighter">🧠 Smart Build</h2>
              <p className="text-sm text-slate-500 font-medium">Hệ thống sẽ bốc ngẫu nhiên câu hỏi theo yêu cầu của bạn.</p>
            </div>

            <div className="space-y-6">
              <div className="space-y-2">
                <label className="text-xs font-black uppercase text-slate-400 tracking-widest ml-1">Nhóm câu hỏi</label>
                <select
                  value={smartConfig.categoryId}
                  onChange={e => setSmartConfig({ ...smartConfig, categoryId: e.target.value })}
                  className="w-full h-14 px-5 bg-slate-50 border-2 border-slate-100 rounded-2xl font-bold focus:border-green-500 outline-none transition-all"
                >
                  <option value="all">Tất cả nhóm</option>
                  {bankCategories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
              </div>

              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-green-600 tracking-widest ml-1">Dễ</label>
                  <input
                    type="number"
                    min="0"
                    value={smartConfig.easy}
                    onChange={e => setSmartConfig({ ...smartConfig, easy: Number(e.target.value) })}
                    className="w-full h-14 px-4 bg-green-50 border-2 border-green-100 rounded-2xl font-black text-center focus:border-green-500 outline-none"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-orange-600 tracking-widest ml-1">T.Bình</label>
                  <input
                    type="number"
                    min="0"
                    value={smartConfig.medium}
                    onChange={e => setSmartConfig({ ...smartConfig, medium: Number(e.target.value) })}
                    className="w-full h-14 px-4 bg-orange-50 border-2 border-orange-100 rounded-2xl font-black text-center focus:border-orange-500 outline-none"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-red-600 tracking-widest ml-1">Khó</label>
                  <input
                    type="number"
                    min="0"
                    value={smartConfig.hard}
                    onChange={e => setSmartConfig({ ...smartConfig, hard: Number(e.target.value) })}
                    className="w-full h-14 px-4 bg-red-50 border-2 border-red-100 rounded-2xl font-black text-center focus:border-red-500 outline-none"
                  />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-3 pt-4">
              <button onClick={() => setIsSmartModalOpen(false)} className="px-6 py-3 rounded-2xl font-bold text-slate-500 hover:bg-slate-50 transition-all">HỦY</button>
              <button
                onClick={handleSmartBuild}
                className="px-8 py-3 rounded-2xl bg-black text-white font-black hover:bg-slate-800 transition-all shadow-xl active:scale-95 uppercase tracking-tight"
              >
                Bốc câu hỏi ngay
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-2">
      <label className="text-sm font-semibold text-gray-700 ml-1">{label}</label>
      {children}
    </div>
  )
}

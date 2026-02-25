'use client'

import { useEffect, useMemo, useState } from 'react'
import { useParams, useRouter } from 'next/navigation'
import { createClient } from '@/utils/supabase/client'

type QuestionType = 'single' | 'multiple' | 'essay'

type Answer = {
  id: string
  content: string
  is_correct: boolean
  images: string[]
}

type Question = {
  id: string
  content: string
  type: QuestionType
  images: string[]
  answers: Answer[]
}

type Submission = {
  id: string
  test_id: string
  candidate_name: string | null
  access_code_id: string
  score_percent: number
  correct_count: number
  total_count: number
  passed: boolean
  started_at: string | null
  submitted_at: string | null
  duration_seconds: number | null
  violation_count: number
  created_at: string
}

type SubmissionAnswer = {
  id: string
  submission_id: string
  question_id: string
  selected_answer_ids: string[] | null
  essay_text: string | null
  is_correct: boolean | null
}

type ViolationLog = {
  id: string
  violation_reason: string
  violated_at: string
  created_at: string
}

function parseImages(val: string | null | undefined): string[] {
  if (!val) return []
  try {
    if (val.trim().startsWith('[')) {
      const parsed = JSON.parse(val)
      if (Array.isArray(parsed)) return parsed.filter((x: any) => typeof x === 'string')
    }
  } catch (e) { }
  return [val]
}

function formatDuration(seconds: number | null) {
  if (!seconds && seconds !== 0) return '-'
  const m = Math.floor(seconds / 60)
  const s = seconds % 60
  return `${m}m ${s}s`
}

function formatDateTime(isoString: string | null) {
  if (!isoString) return '-'
  const date = new Date(isoString)
  const pad = (n: number) => n.toString().padStart(2, '0')
  const day = pad(date.getDate())
  const month = pad(date.getMonth() + 1)
  const year = date.getFullYear()
  const hours = pad(date.getHours())
  const minutes = pad(date.getMinutes())
  const seconds = pad(date.getSeconds())
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`
}

// ────────────────────────────────────────────────────────────
// Export helpers
// ────────────────────────────────────────────────────────────

function buildExportRows(
  submission: Submission,
  questions: Question[],
  subAnswerMap: Map<string, SubmissionAnswer>,
  testTitle: string,
) {
  const rows: Record<string, string>[] = []

  questions.forEach((q, idx) => {
    const sa = subAnswerMap.get(q.id)
    const selectedIds = sa?.selected_answer_ids ?? []
    const correctIds = q.answers.filter(a => a.is_correct).map(a => a.id)

    const selectedAnswers = q.answers
      .filter(a => selectedIds.includes(a.id))
      .map(a => a.content || '[Đáp án hình ảnh]')
      .join(' | ')

    const correctAnswers = q.answers
      .filter(a => a.is_correct)
      .map(a => a.content || '[Đáp án hình ảnh]')
      .join(' | ')

    rows.push({
      'Bài thi': testTitle,
      'Thí sinh': submission.candidate_name?.trim() || 'Ẩn danh',
      'Điểm (%)': String(submission.score_percent),
      'Kết quả': submission.passed ? 'ĐẠT' : 'CHƯA ĐẠT',
      'Thời gian làm': formatDuration(submission.duration_seconds),
      'Số vi phạm': String(submission.violation_count ?? 0),
      'Ngày nộp': formatDateTime(submission.submitted_at ?? submission.created_at),
      'STT câu': String(idx + 1),
      'Nội dung câu hỏi': q.content,
      'Loại câu': q.type === 'essay' ? 'Tự luận' : q.type === 'multiple' ? 'Nhiều đáp án' : 'Một đáp án',
      'Đáp án được chọn': q.type === 'essay' ? (sa?.essay_text ?? '') : selectedAnswers,
      'Đáp án đúng': q.type === 'essay' ? '' : correctAnswers,
      'Đúng/Sai': q.type === 'essay' ? 'Tự luận' : sa?.is_correct === true ? 'Đúng' : 'Sai',
    })
  })

  return rows
}

function exportCSV(rows: Record<string, string>[], filename: string) {
  if (rows.length === 0) return
  const headers = Object.keys(rows[0])
  const escape = (v: string) => `"${(v ?? '').replace(/"/g, '""')}"`
  const lines = [
    headers.map(escape).join(','),
    ...rows.map(r => headers.map(h => escape(r[h] ?? '')).join(','))
  ]
  const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function exportExcel(rows: Record<string, string>[], filename: string) {
  if (rows.length === 0) return
  const headers = Object.keys(rows[0])

  // Build a simple HTML table that Excel can open
  const table = `<table><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${headers.map(h => `<td>${(r[h] ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>`).join('')}</tr>`).join('')
    }</tbody></table>`

  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>
<body>${table}</body></html>`

  const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function copyForGoogleSheet(rows: Record<string, string>[]) {
  if (rows.length === 0) return
  const headers = Object.keys(rows[0])
  const lines = [
    headers.join('\t'),
    ...rows.map(r => headers.map(h => (r[h] ?? '').replace(/\t/g, ' ').replace(/\n/g, ' ')).join('\t'))
  ]
  navigator.clipboard.writeText(lines.join('\n')).then(() => {
    alert('✅ Đã copy dữ liệu! Mở Google Sheet và nhấn Ctrl+V (hoặc Cmd+V) để dán.')
  }).catch(() => {
    // fallback
    const ta = document.createElement('textarea')
    ta.value = lines.join('\n')
    document.body.appendChild(ta)
    ta.select()
    document.execCommand('copy')
    document.body.removeChild(ta)
    alert('✅ Đã copy dữ liệu! Mở Google Sheet và nhấn Ctrl+V (hoặc Cmd+V) để dán.')
  })
}

// ────────────────────────────────────────────────────────────

export default function ReportDetailPage() {
  const supabase = createClient()
  const params = useParams<{ id: string }>()
  const router = useRouter()
  const submissionId = params?.id

  const [loading, setLoading] = useState(true)
  const [submission, setSubmission] = useState<Submission | null>(null)
  const [testTitle, setTestTitle] = useState('')
  const [questions, setQuestions] = useState<Question[]>([])
  const [subAnswers, setSubAnswers] = useState<SubmissionAnswer[]>([])
  const [violationLogs, setViolationLogs] = useState<ViolationLog[]>([])
  const [error, setError] = useState<string | null>(null)
  const [exportMenuOpen, setExportMenuOpen] = useState(false)

  useEffect(() => {
    const run = async () => {
      if (!submissionId) return
      setLoading(true)
      setError(null)

      const { data: sub, error: subErr } = await supabase
        .from('test_submissions')
        .select('*')
        .eq('id', submissionId)
        .single()

      if (subErr || !sub) {
        setError(subErr?.message ?? 'Không tìm thấy submission')
        setLoading(false)
        return
      }
      setSubmission(sub)

      // Load test title
      const { data: testRow } = await supabase
        .from('tests')
        .select('title')
        .eq('id', sub.test_id)
        .single()
      setTestTitle(testRow?.title ?? '')

      const { data: sa, error: saErr } = await supabase
        .from('test_submission_answers')
        .select('*')
        .eq('submission_id', submissionId)

      if (saErr) {
        setError(saErr.message)
        setLoading(false)
        return
      }
      setSubAnswers(sa ?? [])

      // ✅ FIX: Fetch image_url for questions
      const { data: qsRaw, error: qsErr } = await supabase
        .from('questions')
        .select('id, content, type, image_url')
        .eq('test_id', sub.test_id)

      if (qsErr) {
        setError(qsErr.message)
        setLoading(false)
        return
      }

      const qs = (qsRaw ?? []) as Array<{ id: string; content: string; type: QuestionType; image_url: string | null }>
      const questionIds = qs.map(q => q.id)

      const answersByQuestion = new Map<string, Answer[]>()
      if (questionIds.length > 0) {
        // ✅ FIX: Also fetch image_url for answers
        const { data: ansRaw, error: ansErr } = await supabase
          .from('answers')
          .select('id, question_id, content, is_correct, image_url')
          .in('question_id', questionIds)

        if (ansErr) {
          setError(ansErr.message)
          setLoading(false)
          return
        }

        for (const a of (ansRaw ?? []) as any[]) {
          const arr = answersByQuestion.get(a.question_id) ?? []
          arr.push({
            id: a.id,
            content: a.content,
            is_correct: a.is_correct,
            images: parseImages(a.image_url),
          })
          answersByQuestion.set(a.question_id, arr)
        }
      }

      const merged: Question[] = qs.map(q => ({
        id: q.id,
        content: q.content,
        type: q.type,
        images: parseImages(q.image_url),
        answers: answersByQuestion.get(q.id) ?? [],
      }))

      setQuestions(merged)

      // ✅ Fetch violation logs
      const { data: logs, error: logsErr } = await supabase
        .from('test_violation_logs')
        .select('id, violation_reason, violated_at, created_at')
        .eq('access_code_id', sub.access_code_id)
        .order('violated_at', { ascending: true })

      if (!logsErr && logs) {
        setViolationLogs(logs)
      }

      setLoading(false)
    }

    run()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [submissionId])

  const subAnswerMap = useMemo(() => {
    const m = new Map<string, SubmissionAnswer>()
    for (const a of subAnswers) m.set(a.question_id, a)
    return m
  }, [subAnswers])

  const exportRows = useMemo(() => {
    if (!submission) return []
    return buildExportRows(submission, questions, subAnswerMap, testTitle)
  }, [submission, questions, subAnswerMap, testTitle])

  const safeFilename = (testTitle || 'bai-lam').replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s-]/g, '').replace(/\s+/g, '-').slice(0, 50)

  if (loading) return <div className="p-6 text-gray-600">Đang tải...</div>
  if (error) return <div className="p-6 text-red-600">Lỗi: {error}</div>
  if (!submission) return <div className="p-6 text-gray-600">Không có dữ liệu.</div>

  return (
    <div className="p-6 bg-white text-gray-900">
      {/* Header */}
      <div className="flex items-start justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-[var(--primary)]">Chi tiết bài làm</h1>
          {testTitle && (
            <div className="mt-1 text-sm font-semibold text-gray-500">{testTitle}</div>
          )}

          <div className="mt-2 text-sm text-gray-600 space-y-1">
            <div>
              Thí sinh:{' '}
              <span className="font-semibold text-gray-900">
                {submission.candidate_name?.trim() || '(Chưa có tên)'}
              </span>
            </div>

            <div>
              Kết quả:{' '}
              <span className={`font-semibold ${submission.passed ? 'text-green-700' : 'text-red-700'}`}>
                {submission.passed ? 'ĐẠT' : 'CHƯA ĐẠT'}
              </span>
            </div>

            <div>
              Điểm:{' '}
              <span className="font-semibold">
                {submission.score_percent}%
              </span>{' '}
              ({submission.correct_count}/{submission.total_count})
            </div>

            <div>
              Thời gian làm: {formatDuration(submission.duration_seconds)}
            </div>

            <div>
              Số vi phạm:
              <span className={`font-semibold ml-1 ${submission.violation_count > 0 ? 'text-red-700' : 'text-gray-900'}`}>
                {submission.violation_count ?? 0}
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap justify-end">
          {/* Export Button Group */}
          <div className="relative">
            <button
              onClick={() => setExportMenuOpen(v => !v)}
              className="flex items-center gap-2 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-sm shadow-sm transition-all"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5m0 0l5-5m-5 5V4" />
              </svg>
              Xuất dữ liệu
              <svg xmlns="http://www.w3.org/2000/svg" className={`w-3 h-3 transition-transform ${exportMenuOpen ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={3}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
              </svg>
            </button>

            {exportMenuOpen && (
              <div
                className="absolute right-0 mt-1 w-56 bg-white border border-gray-200 rounded-xl shadow-xl z-50 overflow-hidden"
                onMouseLeave={() => setExportMenuOpen(false)}
              >
                <button
                  onClick={() => {
                    exportCSV(exportRows, `${safeFilename}.csv`)
                    setExportMenuOpen(false)
                  }}
                  className="w-full text-left px-4 py-3 text-sm font-semibold text-gray-700 hover:bg-gray-50 flex items-center gap-3 border-b border-gray-100 transition-colors"
                >
                  <span className="text-lg">📄</span>
                  <div>
                    <div>Xuất CSV</div>
                    <div className="text-xs font-normal text-gray-400">Mở bằng Excel / Numbers</div>
                  </div>
                </button>

                <button
                  onClick={() => {
                    exportExcel(exportRows, `${safeFilename}.xls`)
                    setExportMenuOpen(false)
                  }}
                  className="w-full text-left px-4 py-3 text-sm font-semibold text-gray-700 hover:bg-gray-50 flex items-center gap-3 border-b border-gray-100 transition-colors"
                >
                  <span className="text-lg">📊</span>
                  <div>
                    <div>Xuất Excel (.xls)</div>
                    <div className="text-xs font-normal text-gray-400">Microsoft Excel</div>
                  </div>
                </button>

                <button
                  onClick={() => {
                    copyForGoogleSheet(exportRows)
                    setExportMenuOpen(false)
                  }}
                  className="w-full text-left px-4 py-3 text-sm font-semibold text-gray-700 hover:bg-gray-50 flex items-center gap-3 transition-colors"
                >
                  <span className="text-lg">🟢</span>
                  <div>
                    <div>Copy cho Google Sheet</div>
                    <div className="text-xs font-normal text-gray-400">Dán trực tiếp vào Google Sheets</div>
                  </div>
                </button>
              </div>
            )}
          </div>

          <button
            onClick={() => router.push('/reports')}
            className="px-3 py-2 rounded-lg border border-gray-200 hover:bg-gray-50 text-gray-700"
          >
            ← Quay lại
          </button>
        </div>
      </div>

      {/* ✅ Lịch sử Vi phạm */}
      {violationLogs.length > 0 && (
        <div className="mt-6 rounded-2xl border border-red-200 bg-red-50/30 p-5">
          <h2 className="text-lg font-bold text-red-700 mb-3">
            ⚠️ Lịch sử Vi phạm ({violationLogs.length})
          </h2>
          <div className="space-y-2">
            {violationLogs.map((log, idx) => (
              <div
                key={log.id}
                className="rounded-lg border border-red-200 bg-white p-3 flex items-start justify-between gap-3"
              >
                <div className="flex-1">
                  <div className="text-sm font-semibold text-red-700">
                    #{idx + 1}: {log.violation_reason}
                  </div>
                  <div className="text-xs text-gray-500 mt-1">
                    Thời gian: {formatDateTime(log.violated_at)}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Questions */}
      <div className="mt-6 space-y-5">
        {questions.map((q, idx) => {
          const sa = subAnswerMap.get(q.id)
          const selectedIds = sa?.selected_answer_ids ?? []
          const correctIds = q.answers?.filter(a => a.is_correct).map(a => a.id) ?? []
          const isEssay = q.type === 'essay'

          return (
            <div key={q.id} className="rounded-2xl border border-gray-200 bg-white shadow-sm p-5 space-y-3">
              <div className="flex items-start justify-between gap-3">
                <div className="font-semibold">
                  Câu {idx + 1}:{' '}
                  <span className="font-normal text-gray-800">{q.content}</span>
                </div>

                {!isEssay && (
                  <span
                    className={`shrink-0 text-xs px-2.5 py-1 rounded-full border ${sa?.is_correct
                      ? 'border-green-300 bg-green-50 text-green-700'
                      : 'border-red-300 bg-red-50 text-red-700'
                      }`}
                  >
                    {sa?.is_correct ? 'Đúng' : 'Sai'}
                  </span>
                )}

                {isEssay && (
                  <span className="shrink-0 text-xs px-2.5 py-1 rounded-full border border-gray-200 bg-gray-50 text-gray-600">
                    Tự luận
                  </span>
                )}
              </div>

              {/* ✅ FIX: Hiển thị ảnh câu hỏi */}
              {q.images.length > 0 && (
                <div className="flex flex-col gap-3">
                  {q.images.map((img, i) => (
                    <img
                      key={i}
                      src={img}
                      alt={`Hình câu ${idx + 1}`}
                      className="max-w-full max-h-80 rounded-xl border border-gray-200 object-contain bg-gray-50"
                    />
                  ))}
                </div>
              )}

              {/* MCQ */}
              {!isEssay && (
                <div className="space-y-2">
                  {q.answers.map(a => {
                    const isCorrect = correctIds.includes(a.id)
                    const isSelected = selectedIds.includes(a.id)

                    // ✅ Rule:
                    // - Đáp án đúng: xanh
                    // - Chọn sai: đỏ
                    // - Còn lại: xám
                    const cls = isCorrect
                      ? 'border-green-300 bg-green-50'
                      : isSelected
                        ? 'border-red-300 bg-red-50'
                        : 'border-gray-200 bg-white'

                    return (
                      <div key={a.id} className={`rounded-lg border p-3 ${cls}`}>
                        <div className="flex items-start justify-between gap-2">
                          <div className="flex items-start gap-2 flex-1">
                            {/* ✅ FIX: Hiển thị checkbox/radio state chính xác */}
                            <span className={`mt-0.5 shrink-0 w-4 h-4 rounded-full border-2 flex items-center justify-center ${isSelected
                                ? isCorrect
                                  ? 'border-green-500 bg-green-500'
                                  : 'border-red-500 bg-red-500'
                                : isCorrect
                                  ? 'border-green-400 bg-white'
                                  : 'border-gray-300 bg-white'
                              }`}>
                              {isSelected && (
                                <span className="w-2 h-2 rounded-full bg-white block" />
                              )}
                            </span>

                            <div className="flex-1">
                              {/* ✅ Hiển thị text đáp án */}
                              {a.content && (
                                <div className="text-sm text-gray-900">{a.content}</div>
                              )}

                              {/* ✅ FIX: Hiển thị ảnh đáp án với kích thước full (không bị thu nhỏ) */}
                              {a.images.length > 0 && (
                                <div className="mt-2 flex flex-col gap-2">
                                  {a.images.map((img, i) => (
                                    <img
                                      key={i}
                                      src={img}
                                      alt={`Đáp án ${String.fromCharCode(65 + q.answers.indexOf(a))}`}
                                      className="max-w-full max-h-64 rounded-lg border object-contain bg-gray-50"
                                    />
                                  ))}
                                </div>
                              )}
                            </div>
                          </div>

                          <div className="flex gap-2 shrink-0">
                            {isCorrect && (
                              <span className="text-xs px-2 py-1 rounded-full border border-green-300 bg-white text-green-700 whitespace-nowrap">
                                Đáp án đúng
                              </span>
                            )}

                            {isSelected && (
                              <span
                                className={`text-xs px-2 py-1 rounded-full border bg-white whitespace-nowrap ${isCorrect ? 'border-green-300 text-green-700' : 'border-red-300 text-red-700'
                                  }`}
                              >
                                Bạn chọn
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                    )
                  })}

                  {q.answers.length === 0 && (
                    <div className="text-sm text-gray-500 italic">(Không load được đáp án)</div>
                  )}
                </div>
              )}

              {/* Essay */}
              {isEssay && (
                <div className="space-y-2">
                  <div className="text-sm text-gray-600">Bài làm:</div>
                  <div className="rounded-lg border border-gray-200 bg-gray-50 p-3 whitespace-pre-wrap text-sm text-gray-900">
                    {sa?.essay_text?.trim() ? sa.essay_text : <span className="text-gray-400">(Không có nội dung)</span>}
                  </div>

                  <div className="text-xs text-gray-500">
                    Trạng thái chấm: {sa?.is_correct === true ? 'Đúng' : sa?.is_correct === false ? 'Sai' : 'Chưa chấm'}
                  </div>
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* Bottom export bar */}
      {questions.length > 0 && (
        <div className="mt-8 pt-6 border-t border-gray-100 flex flex-wrap items-center justify-between gap-3">
          <div className="text-sm text-gray-500">
            Tổng {questions.length} câu hỏi • {submission.correct_count}/{submission.total_count} câu đúng
          </div>
          <div className="flex gap-2 flex-wrap">
            <button
              onClick={() => exportCSV(exportRows, `${safeFilename}.csv`)}
              className="px-4 py-2 rounded-lg border border-gray-200 bg-white hover:bg-gray-50 text-gray-700 font-semibold text-sm flex items-center gap-2 transition-colors shadow-sm"
            >
              📄 CSV
            </button>
            <button
              onClick={() => exportExcel(exportRows, `${safeFilename}.xls`)}
              className="px-4 py-2 rounded-lg border border-gray-200 bg-white hover:bg-gray-50 text-gray-700 font-semibold text-sm flex items-center gap-2 transition-colors shadow-sm"
            >
              📊 Excel
            </button>
            <button
              onClick={() => copyForGoogleSheet(exportRows)}
              className="px-4 py-2 rounded-lg border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-semibold text-sm flex items-center gap-2 transition-colors shadow-sm"
            >
              🟢 Copy → Google Sheets
            </button>
          </div>
        </div>
      )}
    </div>
  )
}

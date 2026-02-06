-- Add other missing columns to tests table
-- Fixes "Could not find column in schema cache" errors

-- 1. Add shuffle_questions (Support randomizing question order)
ALTER TABLE tests 
ADD COLUMN IF NOT EXISTS shuffle_questions BOOLEAN DEFAULT FALSE;

COMMENT ON COLUMN tests.shuffle_questions IS 'Nếu true, thứ tự các câu hỏi sẽ bị xáo trộn';

-- 2. Add max_violations (Config fail limit for anti-cheat)
ALTER TABLE tests 
ADD COLUMN IF NOT EXISTS max_violations INTEGER DEFAULT 0;

COMMENT ON COLUMN tests.max_violations IS 'Số lần vi phạm tối đa cho phép (0 = không giới hạn). Nếu vượt quá, bài thi sẽ bị khóa.';

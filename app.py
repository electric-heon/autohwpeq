from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import sys
sys.path.append('..')
import pyhwpx
import time
import threading
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 진행 상황 저장용 딕셔너리
progress_data = {}

class HwpEquationAutomation:
    def __init__(self, visible=False, job_id=None):
        self.hwp = pyhwpx.Hwp(visible=visible)
        self.error_log = []
        self.job_id = job_id

    def parse_math_text(self, text):
        """$로 구분된 텍스트에서 수식과 일반 텍스트를 분리"""
        parts = []
        segments = text.split('$')

        for i, segment in enumerate(segments):
            if segment:
                if i % 2 == 0:  # 짝수 인덱스 = 일반 텍스트
                    parts.append(('text', segment))
                else:  # 홀수 인덱스 = 수식
                    parts.append(('math', segment))

        return parts

    def insert_text(self, text):
        """일반 텍스트 삽입"""
        self.hwp.insert_text(text)

    def insert_equation(self, formula):
        """안전한 수식 삽입 (재시도 로직 포함)"""
        max_retries = 3

        for attempt in range(max_retries):
            try:
                # 수식 개체 삽입
                pset = self.hwp.HParameterSet.HEqEdit
                self.hwp.HAction.GetDefault("EquationCreate", pset.HSet)
                pset.string = formula
                pset.BaseUnit = 900
                self.hwp.HAction.Execute("EquationCreate", pset.HSet)

                # 수식 편집 종료
                self.hwp.HAction.Run("Cancel")
                time.sleep(0.05)

                return True

            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(0.2)
                else:
                    error_msg = f"수식 삽입 실패: {formula[:50]}... - {e}"
                    self.error_log.append(error_msg)
                    self.insert_text(f"[수식 오류: {formula[:30]}...]")
                    return False

    def update_progress(self, current, total, message):
        """진행 상황 업데이트"""
        if self.job_id:
            progress_data[self.job_id] = {
                'current': current,
                'total': total,
                'percentage': (current / total) * 100 if total > 0 else 0,
                'message': message,
                'timestamp': datetime.now().isoformat()
            }

    def process_line(self, line, line_num=None, total_lines=None):
        """한 줄 처리 (진행상황 표시 포함)"""
        if not line.strip():
            self.hwp.HAction.Run("BreakPara")
            return

        # 진행률 업데이트
        if line_num is not None and total_lines is not None:
            message = f"줄 {line_num}/{total_lines}: {line[:60]}..."
            self.update_progress(line_num, total_lines, message)

        parts = self.parse_math_text(line)

        for part_type, content in parts:
            if part_type == 'text':
                self.insert_text(content)
            elif part_type == 'math':
                self.insert_equation(content)

        self.hwp.HAction.Run("BreakPara")

    def process_document(self, text, save_path=None):
        """전체 문서 처리 (배치 처리 모드)"""
        lines = text.split('\n')
        total_lines = len(lines)

        self.update_progress(0, total_lines, "문서 처리 시작")
        start_time = time.time()

        for i, line in enumerate(lines, 1):
            self.process_line(line, i, total_lines)

        elapsed_time = time.time() - start_time

        if save_path:
            self.save_document(save_path)

        self.update_progress(total_lines, total_lines, "처리 완료")

        return {
            'total_lines': total_lines,
            'elapsed_time': elapsed_time,
            'errors': len(self.error_log),
            'error_log': self.error_log
        }

    def save_document(self, filepath):
        """문서 저장"""
        self.hwp.SaveAs(filepath)

    def close(self):
        """한글 종료"""
        self.hwp.Quit()


def process_file_background(input_path, output_path, job_id):
    """백그라운드에서 파일 처리"""
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            text = f.read()

        automation = HwpEquationAutomation(visible=False, job_id=job_id)
        result = automation.process_document(text, save_path=output_path)
        automation.close()

        progress_data[job_id]['status'] = 'completed'
        progress_data[job_id]['result'] = result

    except Exception as e:
        progress_data[job_id]['status'] = 'error'
        progress_data[job_id]['error'] = str(e)


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """파일 업로드 및 처리 시작"""
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '파일이 선택되지 않았습니다'}), 400

    if file:
        # 파일 저장
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        output_filename = f"{timestamp}_output.hwp"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        file.save(input_path)

        # Job ID 생성
        job_id = f"job_{timestamp}"

        # 백그라운드에서 처리
        thread = threading.Thread(
            target=process_file_background,
            args=(input_path, output_path, job_id)
        )
        thread.start()

        return jsonify({
            'job_id': job_id,
            'message': '처리가 시작되었습니다'
        })


@app.route('/progress/<job_id>')
def get_progress(job_id):
    """진행 상황 조회"""
    if job_id in progress_data:
        return jsonify(progress_data[job_id])
    else:
        return jsonify({'error': '작업을 찾을 수 없습니다'}), 404


@app.route('/download/<job_id>')
def download_file(job_id):
    """결과 파일 다운로드"""
    if job_id not in progress_data:
        return jsonify({'error': '작업을 찾을 수 없습니다'}), 404

    if progress_data[job_id].get('status') != 'completed':
        return jsonify({'error': '처리가 완료되지 않았습니다'}), 400

    # 출력 파일 찾기
    timestamp = job_id.replace('job_', '')
    output_filename = f"{timestamp}_output.hwp"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    if os.path.exists(output_path):
        return send_file(output_path, as_attachment=True, download_name='result.hwp')
    else:
        return jsonify({'error': '파일을 찾을 수 없습니다'}), 404


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

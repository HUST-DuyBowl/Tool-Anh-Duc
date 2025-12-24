# chatgpt.py
from openai import OpenAI
import json

def phantichfile(apikey, pdfpath, vector_id):
    client = OpenAI(api_key=apikey)
    
    # Sử dụng 'with open' để đảm bảo file được tự động đóng lại an toàn
    with open(pdfpath, "rb") as pdf_file_handle:
        file = client.files.create(
            file=pdf_file_handle,
            purpose="user_data"
        )
    
    prompt = [
        {
          "role": "system",
          "content": [
            {
              "type": "input_text",
              "text": "Bạn là một giáo viên đang tổng hợp thông tin về các đề Toán của mình. Sử dụng dữ liệu được cho, hãy phân tích và cho ra kết quả theo định dạng JSON, không cần giải thích,nếu không có thì bỏ trống:\n-  Loại kỳ thi (giữa học kỳ I, giữa học kỳ II, Khảo sát chất lượng tháng 8, thi thử THPT, thi thử tuyển sinh vào 10, ... sử dụng viết tắt và thêm thông tin như CK II, KSCL tháng 8)\n- Năm học (2023-2024, 2024-2025)\n- Lớp (viết 1 số duy nhất)\n- Chương học (tìm kiếm trong dữ liệu sách được cung cấp)\n- Bài học (tìm kiếm trong dữ liệu sách được cung cấp)\nMột số cụm viết tắt cần thiết: Khảo sát chất lượng: KSCL; giữa kỳ: GK; cuối kỳ: CK"
            }
          ]
        },
        {
            "role": "user",
            "content": [
                {
                    "type": "input_file",
                    "file_id": file.id
                }
            ]
        }
      ]
      
    response = client.responses.create(
    model="gpt-4.1", # Lưu ý: Hãy đảm bảo bạn có quyền truy cập model này, nếu lỗi hãy đổi thành "gpt-4o"
    input=prompt,
    text={
            "format": {
                "type": "json_schema",
                "name": "exam_info",
                "strict": True,
                "schema": {
                    "type": "object",
                    "properties": {
                        "Loại kỳ thi": {
                            "type": "string",
                            "description": "Type of the exam (e.g., 'CK II', 'GK I', 'KSCL tháng 8')"
                        },
                        "Năm học": {
                            "type": "string",
                            "description": "Academic year (e.g., '2023-2024', '2024-2025')"
                        },
                        "Lớp": {
                            "type": "integer",
                            "description": "Class represented by a single number, e.g., 7, 8, 12"
                        },
                        "Chương học": {
                            "type": "string",
                            "description": "Curriculum searched from provided textbook data, only gives the number of the curriculum"
                        },
                        "Bài học": {
                            "type": "string",
                            "description": "Lesson searched from provided textbook data, only gives the number of the lesson"
                        }
                    },
                    "required": [
                        "Loại kỳ thi",
                        "Năm học",
                        "Lớp",
                        "Chương học",
                        "Bài học"
                    ],
                    "additionalProperties": False
                }
            }
        },
    reasoning={},
    tools=[
        {
          "type": "file_search",
          "vector_store_ids": [
            vector_id  # <--- ĐÃ SỬA: Dùng biến truyền vào thay vì mã cứng
          ]
        }
    ],
    tool_choice="file_search",
    temperature=0,
    max_output_tokens=4096,
    top_p=0.5,
    store=True
    )
    
    # Xử lý kết quả trả về
    output = json.loads(response.model_dump_json())
    output = output["output"]
    for i in range(len(output)):
        try:
            if output[i]['type'] == 'message':
                output = output[i]["content"][0]["text"]
        except TypeError:
            pass
    thongtin = json.loads(output)
    #truonghoc = thongtin["Tên trường học"]
    #quanhuyen = thongtin["Quận/huyện"]
    lop = thongtin["Lớp"]
    #bosach = thongtin["Bộ sách"]
    chuonghoc = thongtin["Chương học"]
    baihoc = thongtin["Bài học"]
    namhoc = thongtin["Năm học"]
    loaikythi = thongtin["Loại kỳ thi"]
    return lop, chuonghoc, baihoc, namhoc, loaikythi
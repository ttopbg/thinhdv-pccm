# Chuyển đổi dữ liệu giáo viên

Công cụ chuyển đổi file Excel chứa dữ liệu giáo viên (sheet **Data**) sang định dạng chuẩn gồm 3 sheet: **Class**, **Teachers**, **Students**.

## Cấu trúc file

| File | Mô tả |
|------|-------|
| `teacher_core.py` | Logic xử lý dùng chung |
| `app.py` | Giao diện web (Streamlit) |
| `convert_teachers_local.py` | Giao diện local (tkinter) |
| `requirements.txt` | Thư viện cần thiết |

## Chạy local (tkinter)

```bash
pip install openpyxl pandas anthropic
python convert_teachers_local.py
```

## Deploy lên Streamlit Cloud

1. **Fork / push repo lên GitHub**

2. Vào [https://streamlit.io/cloud](https://streamlit.io/cloud) → **New app**

3. Chọn repo, branch `main`, main file = `app.py` → **Deploy**

4. Sau khi deploy: **Settings → Secrets** → thêm:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```

## Cấu trúc file output

### Sheet `Class` (bên trái)
| Cột A | Cột B |
|-------|-------|
| Niên khóa | 2025-2026 |
| Lớp | Khối |
| 10A1 | 10 |
| 11B2 | 11 |
| … | … |

### Sheet `Teachers`
`STT` · `Họ tên` · `Ngày sinh` · `SĐT` · `Môn dạy` · `TBM` · `CN` · `PCCM`

### Sheet `Students` (bên phải)
`STT` · `Mã HS` · `Họ tên` · `Lớp` · `Giới tính` · `Ngày sinh` · `Số điện thoại` · `Email` · `Tài khoản`

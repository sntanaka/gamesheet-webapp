from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import openpyxl
import tempfile
import os

app = FastAPI()

@app.post("/generate")
async def generate_excel(
    allA: UploadFile = File(...),
    allB: UploadFile = File(...),
    check: UploadFile = File(...)
):
    try:
        # 一時ファイルとして保存
        with tempfile.NamedTemporaryFile(delete=False) as tmpA, \
             tempfile.NamedTemporaryFile(delete=False) as tmpB, \
             tempfile.NamedTemporaryFile(delete=False) as tmpC:
            tmpA.write(await allA.read())
            tmpB.write(await allB.read())
            tmpC.write(await check.read())

        # Excelファイルを読み込む
        wbA = openpyxl.load_workbook(tmpA.name, read_only=True)
        wbB = openpyxl.load_workbook(tmpB.name, read_only=True)
        wbC = openpyxl.load_workbook(tmpC.name, read_only=True)

        # ここにあなたの処理ロジックを追加
        # 例: wbC に wbA と wbB のデータを統合するなど

        output_path = "Gamesheet.xlsx"
        wbC.save(output_path)

        # 一時ファイル削除
        os.remove(tmpA.name)
        os.remove(tmpB.name)
        os.remove(tmpC.name)

        return FileResponse(output_path, filename="Gamesheet.xlsx")

    except Exception as e:
        return {"error": str(e)}

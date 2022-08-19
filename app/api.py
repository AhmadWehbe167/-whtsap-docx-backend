from fastapi import Form
from fastapi import FastAPI, UploadFile
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import re
from pydantic import BaseModel
from docx import Document
from starlette.responses import FileResponse
from docx.enum.text import WD_ALIGN_PARAGRAPH 

class Item(BaseModel):
    start: int
    end: int
    text: list

def convert_to_docx(text):
    document = Document("my-template.docx")
    p=document.add_paragraph(text)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    document.save('generated_file.docx')

# define constants
dir = "app/static"

origins = [
    "*"
]

# Define fastAPI app
app = FastAPI()

# Mount to static folder to be able to access images
# app.mount("/static", StaticFiles(directory=dir), name="static")

# Add CrossMiddleWare to be able to send requests between frontend (react)
# and backend(fastapi) running on different servers
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)


@app.post("/uploadfile")
async def upload_file(file: UploadFile, fileExtension: str = Form(default=True)):
    contents = await file.read()
    contents = contents.decode("utf-8")
    split_pattern = "[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2}\, [0-9]{1,2}\:[0-9]{2} [AP][M] - [^:]*:"
    matches = re.findall(split_pattern, contents)
    text = re.split(split_pattern, contents)
    indices = [i for i in range(len(text)) if "<Media omitted>" not in text[i] and len(text[i]) > 0 ]
    res_text = []
    res_matches = []
    for i in indices[:-1]:
        res_text.append(text[i])
        res_matches.append(matches[i])
    res_text.append( text[indices[-1]])
    return {"matches": res_matches, "text": res_text}

@app.post("/create_file")
async def create_file(data: Item):
    joiner = "\n" + "~"*50 + "\n\n"
    chosen_text = joiner.join(data.text[data.start:data.end-1])
    convert_to_docx(chosen_text)
    return {"status":"Done!"}


@app.get("/generated_file")
async def download_generated_file():
    file_path = "generated_file.docx"
    return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=file_path)
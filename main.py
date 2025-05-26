import gradio as gr
from process import Processor


with gr.Blocks() as app:
    gr.Textbox(label="Обработка Excel файла в Word")
    excel_input = gr.File(label="Загрузите Excel файл (.xlsx)", file_types=[".xlsx"])
    word_input = gr.File(label="Загрузите Word файл (.docx)", file_types=[".docx"])
    process_button = gr.Button("Запустить процесс")
    download_output = gr.File(label="Скачать обработанный файл")

    process_button.click(
        Processor().make_word,
        inputs=[
            word_input,
            excel_input
        ],
        outputs=[download_output],
    )

if __name__ == "__main__":
    app.launch(inbrowser=True)

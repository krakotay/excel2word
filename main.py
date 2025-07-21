import gradio as gr
from process import Processor


with gr.Blocks() as app:
    with gr.Tabs():
        with gr.TabItem("РД_Влияние СПОД"):
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
        with gr.TabItem("РД планирование"):
            origin_input = gr.File(label="Откуда (.xlsx)", file_types=[".xlsx"])
            target_input = gr.File(label="Куда (.xlsx)", file_types=[".xlsx"])
            process_button = gr.Button("Запустить процесс")
            download_output = gr.File(label="Скачать обработанный файл")

            process_button.click(
                Processor().copy_ws,
                inputs=[
                    origin_input,
                    target_input
                ],
                outputs=[download_output],
            )
        with gr.TabItem("Запрос 3"):
            gr.Markdown("""
            `_ВСТАВКА_` - маркер, который будет искаться в ячейках таблиц
            """)
            excel_input = gr.File(label="Загрузите Excel файл (.xlsx)", file_types=[".xlsx"])
            word_input = gr.File(label="Загрузите Word файл (.docx)", file_types=[".docx"])
            process_button = gr.Button("Запустить процесс")
            download_output = gr.File(label="Скачать обработанный файл")

            process_button.click(
                Processor().excel2word_insert,
                inputs=[
                    word_input,
                    excel_input
                ],
                outputs=[download_output],
            )
        with gr.TabItem("Таблицы для отчета"):
            from task_four import insert_tables_with_filter
            gr.Markdown("""
            `_ВСТАВКА_` - маркер, на место которого вставят таблицу
            """)
            excel_input = gr.File(label="Загрузите Excel файл (.xlsx)", file_types=[".xlsx"])
            word_input = gr.File(label="Загрузите Word файл (.docx)", file_types=[".docx"])
            process_button = gr.Button("Запустить процесс")
            download_output = gr.File(label="Скачать обработанный файл")

            process_button.click(
                insert_tables_with_filter,
                inputs=[
                    word_input,
                    excel_input
                ],
                outputs=[download_output],
            )



if __name__ == "__main__":
    app.launch(inbrowser=True)

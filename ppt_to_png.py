import os
import comtypes.client
from pdf2image import convert_from_path


def ppt_to_pdf_png(input_ppt, output_pdf):
    
    foldername = input_ppt.split('.')[0]
    
    if not os.path.exists(foldername):
        os.makedirs(foldername)
    # 啟動 PowerPoint 應用程式
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 2  # 設為 1 代表 PowerPoint 以可見模式運行（0 為不可見）

    # 打開 PPT 檔案
    presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))

    # 另存為 PDF（格式 32 代表 PDF）
    pdf_save_path = os.path.join(foldername, output_pdf)
    #presentation.SaveAs(os.path.abspath(output_pdf), 32)
    presentation.SaveAs(os.path.abspath(pdf_save_path), 32)
    # 關閉 PowerPoint
    presentation.Close()
    powerpoint.Quit()
    
    
    pages = convert_from_path(pdf_save_path, 500)
    for count, page in enumerate(pages):
        page.save(os.path.join(foldername, f'out{count}.jpg'), 'JPEG')
    



ppt_to_pdf_png('t.pptx', 't1.pdf')



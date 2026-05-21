import fitz # PyMuPDF
from django.shortcuts import render
from django.http import HttpResponse


def upload_pdf(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf_file = request.FILES['arquivo']

        # Lê o PDF a partir do arquivo enviado
        doc = fitz.open(stream = pdf_file.read(), filetype = "pdf")

        texto_completo = ""
        for pagina in doc:
            texto_completo += pagina.get_text()

        doc.close()

        # Retorna o arquivo .txt para download
        response = HttpResponse(texto_completo, content_type = 'text/plain')
        response['Content-Disposition'] = 'attachment; filename = "resultado.txt" '
        return response
    return render(request, 'meuapp/upload.html')


from docx import Document

def pdf_para_docx(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf = request.FILES['arquivo']
        doc_fitz = fitz.open(stream=pdf.read(), filetype="pdf")
        
        doc_word = Document()
        for pagina in doc_fitz:
            texto = pagina.get_text()
            doc_word.add_paragraph(texto)
            doc_word.add_page_break()
        
        buffer = BytesIO()
        doc_word.save(buffer)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        response['Content-Disposition'] = 'attachment; filename="resultado.docx"'
        return response
    return render(request, 'meuapp/upload.html')


from openpyxl import Workbook
from io import BytesIO

def pdf_para_xlsx(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf = request.FILES['arquivo']
        doc_fitz = fitz.open(stream=pdf.read(), filetype="pdf")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Conteúdo"
        
        for i, pagina in enumerate(doc_fitz):
            ws.append([f"--- Página {i+1} ---"])
            for linha in pagina.get_text().split('\n'):
                ws.append([linha])
        
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="resultado.xlsx"'
        return response
    return render(request, 'meuapp/upload.html')


from reportlab.pdfgen import canvas
from PIL import Image

def imagem_para_pdf(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        imagem = request.FILES['arquivo']
        img = Image.open(imagem)
        largura, altura = img.size
        
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=(largura, altura))
        
        img_buffer = BytesIO()
        img.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        from reportlab.lib.utils import ImageReader
        c.drawImage(ImageReader(img_buffer), 0, 0, largura, altura)
        c.save()
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="resultado.pdf"'
        return response
    return render(request, 'meuapp/upload.html')


import pikepdf

def mesclar_pdfs(request):
    if request.method == 'POST':
        arquivos = request.FILES.getlist('arquivos')  # múltiplos arquivos
        
        pdf_final = pikepdf.Pdf.new()
        
        for arq in arquivos:
            pdf_temp = pikepdf.Pdf.open(BytesIO(arq.read()))
            pdf_final.pages.extend(pdf_temp.pages)
        
        buffer = BytesIO()
        pdf_final.save(buffer)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="mesclado.pdf"'
        return response
    return render(request, 'meuapp/upload.html')


def dividir_pdf(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf = request.FILES['arquivo']
        pagina_str = request.POST.get('paginas', '')  # ex: "1,3,5"
        
        doc = pikepdf.Pdf.open(BytesIO(pdf.read()))
        paginas = [int(p.strip()) - 1 for p in pagina_str.split(',')]
        
        novo_pdf = pikepdf.Pdf.new()
        for i in paginas:
            if 0 <= i < len(doc.pages):
                novo_pdf.pages.append(doc.pages[i])
        
        buffer = BytesIO()
        novo_pdf.save(buffer)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="dividido.pdf"'
        return response
    return render(request, 'meuapp/upload.html')


def comprimir_pdf(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf = request.FILES['arquivo']
        
        doc = pikepdf.Pdf.open(BytesIO(pdf.read()))
        
        buffer = BytesIO()
        doc.save(buffer, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="comprimido.pdf"'
        return response
    return render(request, 'meuapp/upload.html')


def proteger_pdf(request):
    if request.method == 'POST' and request.FILES.get('arquivo'):
        pdf = request.FILES['arquivo']
        senha = request.POST.get('senha', '123456')
        
        doc = pikepdf.Pdf.open(BytesIO(pdf.read()))
        
        permissoes = pikepdf.Permissions(
            extract=False,
            modify_annotation=False,
            modify_assembly=False,
        )
        
        encriptacao = pikepdf.Encryption(
            user=senha,
            owner=senha + "_owner",
            allow=permissoes
        )
        
        buffer = BytesIO()
        doc.save(buffer, encryption=encriptacao)
        buffer.seek(0)
        
        response = HttpResponse(buffer, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="protegido.pdf"'
        return response
    return render(request, 'meuapp/upload.html')


try:
    import pytesseract
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_DISPONIVEL = True
except ImportError:
    PDF2IMAGE_DISPONIVEL = False

def ocr_pdf(request):
    if not OCR_DISPONIVEL:
        return HttpResponse(
            "OCR não disponível: instale pytesseract.",
            status=501, content_type='text/plain'
        )
    if request.method == 'POST' and request.FILES.get('arquivo'):
        import io
        pdf = request.FILES['arquivo']
        lang = request.POST.get('lang', 'por')
        doc = fitz.open(stream=pdf.read(), filetype="pdf")
        texto_total = ""
        for i, pagina in enumerate(doc):
            # Converte página em imagem usando PyMuPDF (sem Poppler)
            mat = fitz.Matrix(2, 2)  # zoom 2x para melhor OCR
            pix = pagina.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("png")
            from PIL import Image
            img = Image.open(io.BytesIO(img_bytes))
            texto = pytesseract.image_to_string(img, lang=lang)
            texto_total += f"\n--- Página {i+1} ---\n{texto}"
        doc.close()
        response = HttpResponse(texto_total, content_type='text/plain; charset=utf-8')
        response['Content-Disposition'] = 'attachment; filename="ocr_resultado.txt"'
        return response
    return render(request, 'meuapp/upload.html')
import fitz # PyMuPDF
from django.shortcuts import render
from django.http import HttpResponse

def upload_pdf(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']

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

# Create your views here.

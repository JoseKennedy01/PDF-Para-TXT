from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_pdf, name='home'),
    path('pdf-para-docx/', views.pdf_para_docx, name='pdf_docx'),
    path('pdf-para-xlsx/', views.pdf_para_xlsx, name='pdf_xlsx'),
    path('imagem-para-pdf/', views.imagem_para_pdf, name='img_pdf'),
    path('mesclar/', views.mesclar_pdfs, name='mesclar'),
    path('dividir/', views.dividir_pdf, name='dividir'),
    path('comprimir/', views.comprimir_pdf, name='comprimir'),
    path('proteger/', views.proteger_pdf, name='proteger'),
    path('ocr/', views.ocr_pdf, name='ocr'),
]
from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd
from PyPDF2 import PdfReader
import os
import mimetypes


def home(request):
    preview_data = None  # Initialize preview data
    preview_data_html = None  # Store HTML version of preview data
    uploaded_file_path = None  # Store the full file path

    if request.method == 'POST':
        if 'file' in request.FILES:
            # Handle the file upload
            uploaded_file = request.FILES['file']

            # Check file type on the backend (for security)
            mime_type, _ = mimetypes.guess_type(uploaded_file.name)
            if mime_type != 'application/pdf':
                return HttpResponse("Invalid file type. Please upload a PDF file.", status=400)

            # Save the uploaded PDF file temporarily
            uploaded_file_path = f'temp/{uploaded_file.name}'
            os.makedirs('temp', exist_ok=True)  # Ensure temp folder exists
            with open(uploaded_file_path, 'wb+') as destination:
                for chunk in uploaded_file.chunks():
                    destination.write(chunk)

            # Extract the data from the PDF
            try:
                preview_data = extract_pdf_data(uploaded_file_path)  # Extract data from PDF
                preview_data_html = preview_data.head().to_html(
                    classes='table table-striped')  # Convert the first few rows to HTML
            except Exception as e:
                return HttpResponse(f"Error processing file: {e}")

        elif 'convert' in request.POST and 'file_path' in request.POST:
            # Handle the Excel conversion after preview
            uploaded_file_path = request.POST.get('file_path')  # Get the saved file path from hidden input
            preview_data = extract_pdf_data(uploaded_file_path)

            # Convert PDF data to Excel
            excel_path = f'temp/{os.path.basename(uploaded_file_path).split(".")[0]}.xlsx'
            preview_data.to_excel(excel_path, index=False)  # Save as Excel

            # Prepare the Excel file for download
            with open(excel_path, 'rb') as excel_file:
                response = HttpResponse(
                    excel_file.read(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                response['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_path)}"'
                return response

    return render(request, 'home.html',
                  {'preview_data_html': preview_data_html, 'uploaded_file_path': uploaded_file_path})


def extract_pdf_data(file_path):
    """
    Extracts data from the PDF file.
    Each page's text is added to a DataFrame.
    """
    reader = PdfReader(file_path)
    data = []

    # Iterate through each page in the PDF
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text()
        data.append({'Page': page_number, 'Text': text})

    # Convert the extracted data into a DataFrame
    df = pd.DataFrame(data)
    return df


from django.shortcuts import render
from .models import Dsaform, Dsauser, Dsauserform
from .forms import UploadForm, DownloadForm
from django.db.models.signals import post_save
from django.dispatch import receiver
from django.http import HttpResponse, HttpResponseNotFound

FILE_TYPE = ['xlsx']


def index(request):                                 # function to set the view of homepage

    return render(request, 'home/index.html')


def upload(request):                                # function to set the view of upload page
    form = UploadForm(request.POST or None, request.FILES or None)
    if form.is_valid():
        dsaform = form.save(commit=False)
        dsaform.excel_text = request.user                   # stores the user typed filename
        dsaform.excel_file = request.FILES['excel_file']    # stores uploaded file

        file_type = dsaform.excel_file.url.split('.')[-1]           # splits filename and keeps latter i.e. xlsx
        file_type = file_type.lower()
        if file_type not in FILE_TYPE:                      # checks if uploaded file is .xlsx
            context = {
                'dsaform': dsaform,
                'form': form,
                'error_message': 'The file you uploaded must be in .XLSX format',
            }
            return render(request, 'home/upload_form.html', context)
        dsaform.save()
        return render(request, 'home/index.html', {'dsaform': dsaform})
    context = {
        "form": form,
    }
    return render(request, 'home/upload_form.html', context)


def download(request):  # function to set the view of upload page

    form = DownloadForm(request.POST or None, request.FILES or None)
    if form.is_valid():
        teacher_name = form.cleaned_data.get('teacher_name')
        teacher_name = teacher_name.replace(' ', '')
        dsauser = Dsauser.objects.filter(userfile_name=teacher_name.lower())

        if dsauser.exists():
            dsauser_obj = dsauser.first()
            file_location = dsauser_obj.user_file.path
            try:
                with open(file_location, 'rb') as f:
                    document = f.read()

                    # sending response
                response = HttpResponse(document,
                                        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = 'attachment; filename="download.docx"'

            except IOError:
                # handle file not exist case here
                response = HttpResponseNotFound('<h1>File does not exist. Please enter valid username.</h1>')

            return response
        else:
            response = HttpResponseNotFound('<h1>File does not exist. Please enter valid username.</h1>')
            return response

    context = {
        "form": form,
    }
    return render(request, 'home/download_form.html', context)



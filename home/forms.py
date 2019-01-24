from django import forms
from django.contrib.auth.models import User

from .models import Dsaform, Dsauserform




class UploadForm(forms.ModelForm):

    class Meta:
        model = Dsaform
        fields = ['excel_text', 'excel_file']


class DownloadForm(forms.ModelForm):

    class Meta:
        model = Dsauserform
        fields = ['teacher_name']







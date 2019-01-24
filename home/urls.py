from django.conf.urls import url
from . import views

urlpatterns = [
    # /dsasite/
    url(r'^$', views.index, name='index'),


    # /dsasite/upload/
    url(r'^upload/$', views.upload, name='upload'),

    # /dsasite/upload/
    url(r'^download/$', views.download, name='download'),
]

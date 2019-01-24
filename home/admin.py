from django.contrib import admin
from .models import Dsaform
from .models import Dsauser
from .models import Dsauserform

admin.site.register(Dsaform)
admin.site.register(Dsauser)
admin.site.register(Dsauserform)


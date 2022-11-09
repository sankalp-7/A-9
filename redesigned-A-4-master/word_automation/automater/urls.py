from django.urls import path 
from.views import home,download,refresh


app_name='automater'

urlpatterns=[ 
    path('',home,name='home'),
    path('download/',download,name='download'),
    path('refresh/',refresh,name='refresh'),
]
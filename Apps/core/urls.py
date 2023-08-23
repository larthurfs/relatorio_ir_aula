from django.urls import path
from Apps.core.views import home,modelo_excel, criarrelatorio

urlpatterns = [
    path('', home,name='home'),
    path('modelorelatorio', modelo_excel,name='modelorelatorio'),
    path('criarrelatorio', criarrelatorio,name='criarrelatorio'),
]

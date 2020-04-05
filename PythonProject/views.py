from django.shortcuts import render
from PythonProject.papet.papet import *
import requests


def button(request):
    return render(request, 'home.html')


def output(request):
    data = requests.get("https://reqres.in/api/users")
    print(data.text)
    data = data.text
    # test('working')
    return render(request, 'home.html', {'data': data})

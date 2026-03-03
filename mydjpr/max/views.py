from django.http import HttpResponse
import subprocess
from django.shortcuts import render
from max.dopzakazng import dop_ng
from max.grafik import process_transport_data
from max.mak import maks
from max.dopzakaz import dop_shafiev, dop_budyakova, dop_grechushkin, dop_kunavina, dop_torgashina
from max.minparty import minpartyf
from max.nal_po_form import nalichie_po_formatam
from max.nalichie import nalichie_comp, nalichie_rozn, nalichie_comp_RF, run_all_nalichie_analysis
from max.nelikvid import nelikvids
from max.nps import npsview
from max.nps_china import extract_nps, add_to_orders
from max.optzakaz import optzf
from max.optzakazfive import optzakazfivew
from max.nacenka import nacenkaview
from max.block import calculate_eb, calculate_pb, calculate_bn, calculate_rb
from django.http import JsonResponse
import time
import os
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from max.send_letter import send_letter

from max.sku_count import sku_countw




def redirect_to_home_or_login(request):
    if request.user.is_authenticated:
        return redirect('index_page')
    else:
        return redirect('login')


def login_view(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            return redirect('index_page')
        else:
            messages.error(request, 'Неверный логин или пароль')
    return render(request, 'login.html')

def logout_view(request):
    logout(request)
    return redirect('login')



def index_page(request):
    return render(request, 'index.html')


def maksimaln_page(request):
    return render(request, 'maksimaln.html')


def dopzakaz_page(request):
    return render(request, 'dopzakaz.html')


def nacenka_page(request):
    return render(request, 'nacenka.html')

def nalichie_po_form_page(request):
    return render(request, 'nalichie_po_formatam.html')

def nps_china_page(request):
    return render(request, 'nps_china.html')


def dopzakazng_page(request):
    user = request.user
    if user.username in ["admin", "v.grechushkin"]:
        return render(request, 'dopzakazng.html')
    else:
        return render(request, 'dopzakazng1.html')


def optzakaz_page(request):
    return render(request, 'optzakaz.html')

def block_page(request):
    return render(request, 'block.html')

def optzakazfive_page(request):
    return render(request, 'optzakazfive.html')

def nalichie_page(request):
    return render(request, 'nalichie.html')

def nelikvid_page(request):
    return render(request, 'nelikvid.html')

def nps_page(request):
    return render(request, 'nps.html')

def minparty_page(request):
    return render(request, 'minparty.html')

def sku_count_page(request):
    return render(request, 'sku_count.html')

def grafik_page(request):
    return render(request, 'grafik.html')

def sendletter_page(request):
    return render(request, 'sendletter.html')



def maxzapas(request):

    user = request.user
    print(f"User: {user} максзапас")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            maks(file_path)
            return render(request, 'ready.html')
        except Exception as e:
            return render(request, 'error.html', {'error_message': str(e)})
        

def dopzakaz(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} допзаказ")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')

        try:
            if user.username == 't.shafiev' or user.username == 'admin':
                dop_shafiev(file_path, lambda message: context['messages'].append(message))
            elif user.username == 'e.budyakova':
                dop_budyakova(file_path, lambda message: context['messages'].append(message))
            elif user.username == 'v.grechushkin':
                dop_grechushkin(file_path, lambda message: context['messages'].append(message))
            elif user.username == 't.kunavina':
                dop_kunavina(file_path, lambda message: context['messages'].append(message))
            elif user.username == 'k.torgashina':
                dop_torgashina(file_path, lambda message: context['messages'].append(message))
            elif user.username == 's.kretov':
                dop_budyakova(file_path, lambda message: context['messages'].append(message))
            else:
                dop_shafiev(file_path, lambda message: context['messages'].append(message))
        except Exception as e:
            context['messages'].append(str(e))  # Добавляем сообщение об ошибке в messages

    return render(request, 'dopzakaz.html', context)


def dopzakazng(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} допзаказ ng")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')

        try:
            dop_ng(file_path, lambda message: context['messages'].append(message))
        except Exception as e:
            context['messages'].append(str(e))  # Добавляем сообщение об ошибке в messages

    if user.username in ["admin", "v.grechushkin"]:
        return render(request, 'dopzakazng.html', context)
    else:
        return render(request, 'dopzakazng1.html', context)




def block(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} блокировки")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        action = request.POST.get('action')
        print(f"Action: {action}")

        try:
            if action == 'eb':
                calculate_eb(file_path, lambda message: context['messages'].append(message))
            elif action == 'pb':
                calculate_pb(file_path, lambda message: context['messages'].append(message))
            elif action == 'bn':
                calculate_bn(file_path, lambda message: context['messages'].append(message))
            elif action == 'rb':
                calculate_rb(file_path, lambda message: context['messages'].append(message))
        except Exception as e:
            context['messages'].append(str(e))

    return render(request, 'block.html', context)
    


def optzakaz(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} оптзаказ по Юршиной")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            optzf(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'optzakaz.html', context)


def optzakazfive(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} оптзаказ с увеличением до 5тр")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        checkbox_CopyToOrders = request.POST.get('CopyToOrders')
        checkbox_nul_3tr = request.POST.get('nul_3tr')
        checkbox_not_nul_5tr_menee_pol_up = request.POST.get('not_nul_po_pravilam')
        porog_max = int(request.POST.get('max'))
        porog_min = int(request.POST.get('min'))
        try:
            optzakazfivew(file_path, checkbox_CopyToOrders, checkbox_nul_3tr, checkbox_not_nul_5tr_menee_pol_up, porog_max, porog_min, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'optzakazfive.html', context)


def nalichie(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} наличие")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        action = request.POST.get('action')
        print(f"Action: {action}")

        try:
            if action == 'nv':
                run_all_nalichie_analysis(file_path, lambda message: context['messages'].append(message))
            elif action == 'nr':
                nalichie_rozn(file_path, lambda message: context['messages'].append(message))
            elif action == 'nc':
                nalichie_comp(file_path, lambda message: context['messages'].append(message))
            elif action == 'ncrf':
                nalichie_comp_RF(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))

    return render(request, 'nalichie.html', context)


def nelikvid(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} неликвиды")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        porog1 = int(request.POST.get('porog1'))
        porog2 = int(request.POST.get('porog2'))
        porog3 = int(request.POST.get('porog3'))
        porog4 = int(request.POST.get('porog4'))
        porog5 = int(request.POST.get('porog5'))
        porog6 = int(request.POST.get('porog6'))

        try:
            nelikvids(file_path, porog1, porog2, porog3, porog4, porog5, porog6, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))

    return render(request, 'nelikvid.html', context)


def minparty(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} минпартия")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        porog1 = int(request.POST.get('porog1'))
        porog2 = int(request.POST.get('porog2'))
        porog3 = int(request.POST.get('porog3'))
        semena = int(request.POST.get('semena'))
        melk = int(request.POST.get('melk'))
        pod_zakup = int(request.POST.get('pod_zakup'))
        koef_okrugl = float(request.POST.get('koef_okrugl'))

        try:
            minpartyf(file_path, porog1, porog2, porog3, semena, melk, pod_zakup, koef_okrugl, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))

    return render(request, 'minparty.html', context)


def nacenka(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} наценка")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            nacenkaview(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'nacenka.html', context)


def nalichie_po_form(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} наличие по форматам")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            nalichie_po_formatam(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'nalichie_po_formatam.html', context)


def nps(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} NPS")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            npsview(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'nps.html', context)


def nps_china(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} nps_china")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        action = request.POST.get('action')
        print(f"Action: {action}")

        try:
            if action == 'nps_orders':
                extract_nps(file_path, lambda message: context['messages'].append(message))
            elif action == 'add_to_orders_nps':
                add_to_orders(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))

    return render(request, 'nps_china.html', context)
        

def sku_count_view(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} sku_count")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            sku_countw(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'sku_count.html', context)


def grafik_view(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} grafik")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            process_transport_data(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'grafik.html', context)


def sendletter_view(request):
    context = {'messages': []}

    user = request.user
    print(f"User: {user} send_letter")

    if request.method == 'POST':
        file_path = request.POST.get('file_path')
        try:
            send_letter(file_path, lambda message: context['messages'].append(message))
            
        except Exception as e:
            context['messages'].append(str(e))   
            
    return render(request, 'sendletter.html', context)
        


        


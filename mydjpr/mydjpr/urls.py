"""
URL configuration for mydjpr project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import include, path
from django.views.generic import RedirectView
from supply import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.redirect_to_home_or_login, name='redirect'),
    path('home/', include([
        path('', views.index_page, name='index_page'),
        path('maksimaln.html', views.maksimaln_page, name='maksimaln'),
        path('maxzapas/', views.maxzapas, name='maxzapas'),
        path('dopzakaz.html', views.dopzakaz_page, name='dopzakaz_page'),
        path('dopzakaz/', views.dopzakaz, name='dopzakaz'),
        path('dopzakazng.html', views.dopzakazng_page, name='dopzakazng_page'),
        path('dopzakazng/', views.dopzakazng, name='dopzakazng'),
        path('optzakaz.html', views.optzakaz_page, name='optzakaz_page'),
        path('optzakaz/', views.optzakaz, name='optzakaz'),
        path('block.html', views.block_page, name='block_page'),
        path('block/', views.block, name='block'),
        path('optzakazfive.html', views.optzakazfive_page, name='optzakazfive_page'),
        path('optzakazfive/', views.optzakazfive, name='optzakazfive'),
        path('nalichie.html', views.nalichie_page, name='nalichie_page'),
        path('nalichie/', views.nalichie, name='nalichie'),
        path('nelikvid.html', views.nelikvid_page, name='nelikvid_page'),
        path('nelikvid/', views.nelikvid, name='nelikvid'),
        path('minparty.html', views.minparty_page, name='minparty_page'),
        path('minparty/', views.minparty, name='minparty'),
        path('nacenka.html', views.nacenka_page, name='nacenka_page'),
        path('nacenka/', views.nacenka, name='nacenka'),
        path('nalichie_po_formatam.html', views.nalichie_po_form_page, name='nalichie_po_form_page'),
        path('nalichie_po_formatam/', views.nalichie_po_form, name='nalichie_po_form'),
        path('nps.html', views.nps_page, name='nps_page'),
        path('nps/', views.nps, name='nps'),
        path('nps_china.html', views.nps_china_page, name='nps_china_page'),
        path('nps_china/', views.nps_china, name='nps_china'),
        path('grafik.html', views.grafik_page, name='grafik_page'),
        path('grafik/', views.grafik_view, name='grafik_view'),
        path('sku_count.html', views.sku_count_page, name='sku_count_page'),
        path('sku_count/', views.sku_count_view, name='sku_count_view'),
        path('favicon.ico', RedirectView.as_view(url='/static/favicon1.ico')),
    ])),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
]

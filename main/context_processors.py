from .models import Country

def countries(request):
    if request.user.is_authenticated:
        if request.user.is_superuser:
            countries = Country.objects.all()
            if not request.session.get('country'):
                request.session['country'] = Country.objects.all().first().id
                choosen_country = Country.objects.all().first()
            else:
                choosen_country = Country.objects.filter(pk=request.session.get('country')).first()
        else:
            countries = request.user.countries.all()
            # countries = Country.objects.all()
            if int(request.session.get('country')) in countries.values_list('id', flat=True):
                choosen_country = Country.objects.filter(pk=request.session.get('country')).first()
            elif countries.first():
                choosen_country = countries.first()
                request.session['country'] = countries.first().id
            else:
                choosen_country = request.user.countries.all().first()
    else:
        choosen_country = None
        countries = None
    return {'countries': countries,'choosen_country': choosen_country}

def is_arc(request):
    context = {'is_arc':False}
    if request.user.is_authenticated:
        country = Country.objects.get(id=request.session.get('country'))
        if country.name == 'United States':
            context['is_arc'] = True
    return context

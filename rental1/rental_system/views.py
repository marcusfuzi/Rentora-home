from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from .models import Property, Visitor, Owner, Rented, Review, User
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.db import connection, models
from django.db.models import Q
import urllib.request
import json

##########################################################

def dictfetchall(cursor):
    "Return all rows from a cursor as a dict"
    columns = [col[0] for col in cursor.description]
    return [
        dict(zip(columns, row))
        for row in cursor.fetchall()
    ]

def login_user(request):
    error = None
    if request.method == "POST":
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        user = authenticate(username=username, password=password)
        if user is not None:
            login(request, user)
            if user.is_owner:
                return redirect("ownerdashboard")
            else:
                return redirect("dashboard")
        else:
            error = "Invalid credentials. Please check your username and password."
    return render(request, "login.html", {'error': error})

##########################################################

@login_required
def dashboard(request):
    error = None  
    current_user = request.user

    if current_user.is_owner:
        return redirect('ownerdashboard')

    try:
        curr_visitor = Visitor.objects.get(user=current_user)
    except Visitor.DoesNotExist:
        curr_visitor = Visitor.objects.create(user=current_user, profile="Visitor Profile", pref_location="")

    rented_ids = Rented.objects.values_list('prop_id_id', flat=True)
    prop_qs = Property.objects.exclude(id__in=rented_ids)

    # Accept filters via GET or POST
    data = request.GET if request.GET else request.POST

    loc = data.get('location', '').strip()
    prop_type = data.get('property_type', '').strip()
    bhk = data.get('bhk', '').strip()
    min_price = data.get('min_price', '').strip()
    max_price = data.get('max_price', '').strip()
    furnishing = data.get('furnishing', '').strip()
    pref_tenant = data.get('preferred_tenant', '').strip()
    parking = data.get('parking', '').strip()
    amenities_list = data.getlist('amenity')
    sort_by = data.get('sort_by', 'newest').strip()

    # Legacy support
    legacy_val = data.get('val', '').strip()
    if legacy_val and not loc:
        legacy_query = data.get('query', '4')
        if legacy_query == '3':
            loc = legacy_val
        elif legacy_query == '4':
            max_price = legacy_val

    # 1. Location / Keyword Filter
    if loc:
        prop_qs = prop_qs.filter(
            models.Q(location__icontains=loc) |
            models.Q(title__icontains=loc) |
            models.Q(description__icontains=loc)
        )

    # 2. Property Type Filter
    if prop_type and prop_type.lower() != 'all':
        prop_qs = prop_qs.filter(property_type__icontains=prop_type)

    # 3. BHK Filter
    if bhk and bhk.lower() != 'all':
        if bhk.startswith('4'):
            prop_qs = prop_qs.filter(bedrooms__gte=4)
        else:
            try:
                prop_qs = prop_qs.filter(bedrooms=int(bhk))
            except ValueError:
                pass

    # 4. Rent Range Filter
    if min_price:
        try:
            prop_qs = prop_qs.filter(price__gte=int(min_price))
        except ValueError:
            pass
    if max_price:
        try:
            prop_qs = prop_qs.filter(price__lte=int(max_price))
        except ValueError:
            pass

    # 5. Furnishing Filter
    if furnishing and furnishing.lower() != 'all':
        prop_qs = prop_qs.filter(furnishing__icontains=furnishing)

    # 6. Preferred Tenant Filter
    if pref_tenant and pref_tenant.lower() != 'all':
        prop_qs = prop_qs.filter(preferred_tenant__icontains=pref_tenant)

    # 7. Parking Filter
    if parking and parking.lower() != 'all':
        prop_qs = prop_qs.filter(parking__icontains=parking)

    # 8. Amenities Checkboxes Filter
    for am in amenities_list:
        if am.strip():
            prop_qs = prop_qs.filter(amenities__icontains=am.strip())

    # 9. Sorting Order
    if sort_by == 'price_asc' or sort_by == '3':
        prop_qs = prop_qs.order_by('price')
    elif sort_by == 'price_desc':
        prop_qs = prop_qs.order_by('-price')
    elif sort_by in ('rating_desc', '1', '2'):
        prop_qs = prop_qs.order_by('-avg_rating')
    else:
        prop_qs = prop_qs.order_by('-id')

    total_matches = prop_qs.count()

    paginator = Paginator(prop_qs, 9)
    page = request.GET.get('page')
    try:
        property_disp = paginator.page(page)
    except PageNotAnInteger:
        property_disp = paginator.page(1)
    except EmptyPage:
        property_disp = paginator.page(paginator.num_pages)

    my_rented = Rented.objects.filter(visitor_id=curr_visitor).select_related('prop_id', 'owner_id')
    my_rented_props = [r.prop_id for r in my_rented]

    context = {
        'curr_visitor': curr_visitor,
        'properties': property_disp,
        'myprop': my_rented_props,
        'my_rented': my_rented,
        'error': error,
        'total_matches': total_matches,
        'filter_params': {
            'location': loc,
            'property_type': prop_type,
            'bhk': bhk,
            'min_price': min_price,
            'max_price': max_price,
            'furnishing': furnishing,
            'preferred_tenant': pref_tenant,
            'parking': parking,
            'amenities': amenities_list,
            'sort_by': sort_by,
        }
    }
    return render(request, "visitor_dashboard.html", context)

##########################################################

@login_required
def owner_dashboard(request):
    current_user = request.user
    if not current_user.is_owner:
        return redirect('dashboard')

    try:
        current_owner = Owner.objects.get(user=current_user)
    except Owner.DoesNotExist:
        current_owner = Owner.objects.create(user=current_user, owner_name=current_user.username, num_properties=0)

    tenants = Rented.objects.filter(owner_id=current_owner).select_related('prop_id', 'visitor_id')
    properties = Property.objects.filter(owner=current_owner).order_by('-id')

    if request.method == "POST":
        if 'sort' in request.POST:
            if request.POST['sort'] == '1':
                tenants = tenants.order_by('visitor_id__user__username')
            elif request.POST['sort'] == '2':
                tenants = tenants.order_by('rent_to_be_paid')

    paginator = Paginator(properties, 6)
    page = request.GET.get('page')
    try:
        property_disp = paginator.page(page)
    except PageNotAnInteger:
        property_disp = paginator.page(1)
    except EmptyPage:
        property_disp = paginator.page(paginator.num_pages)

    context = {
        'current_owner': current_owner,
        'tenants': tenants,
        'properties': property_disp
    }
    return render(request, "owner_dashboard.html", context)

##########################################################

def register(request):
    error = None
    if request.method == "POST":
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        name = request.POST.get('name', '').strip()
        
        if not username or not password or not name:
            error = "Please fill in all fields."
            return render(request, "register.html", {'error': error, 'username': username, 'name': name})

        existing = User.objects.filter(username__iexact=username)
        if existing.exists():
            error = f"Username '{username}' already exists. Kindly provide another username or login."
            return render(request, "register.html", {'error': error, 'name': name})

        new_user = User.objects.create_user(username=username, password=password, is_owner=True)
        new_user.save()
        new_owner = Owner.objects.create(user=new_user, owner_name=name, num_properties=0)
        new_owner.save()
        
        login(request, new_user)
        return redirect('ownerdashboard')
    return render(request, "register.html", {'error': error})

##########################################################

def register_visitor(request):
    error = None
    if request.method == "POST":
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        profile = request.POST.get('profile', '').strip()
        pref_location = request.POST.get('pref_location', '').strip()

        if not username or not password or not profile or not pref_location:
            error = "Please fill in all fields."
            return render(request, "register_visitor.html", {
                'error': error,
                'username': username,
                'pref_location': pref_location,
                'profile': profile
            })

        existing = User.objects.filter(username__iexact=username)
        if existing.exists():
            error = f"Username '{username}' already exists. Kindly provide another username or login."
            return render(request, "register_visitor.html", {
                'error': error,
                'pref_location': pref_location,
                'profile': profile
            })

        new_user = User.objects.create_user(username=username, password=password, is_owner=False)
        new_user.save()
        new_visitor = Visitor.objects.create(user=new_user, profile=profile, pref_location=pref_location)
        new_visitor.save()

        login(request, new_user)
        return redirect('dashboard')
    return render(request, "register_visitor.html", {'error': error})

##########################################################

@login_required
def view_prop(request, id):
    error = None
    current_user = request.user
    
    current_owner = None
    curr_visitor = None
    if current_user.is_owner:
        try:
            current_owner = Owner.objects.get(user=current_user)
        except Owner.DoesNotExist:
            current_owner = Owner.objects.create(user=current_user, owner_name=current_user.username)
    else:
        try:
            curr_visitor = Visitor.objects.get(user=current_user)
        except Visitor.DoesNotExist:
            curr_visitor = Visitor.objects.create(user=current_user, profile="Visitor Profile")

    try:
        prop = Property.objects.get(id=id)
    except Property.DoesNotExist:
        return redirect('home')

    if not current_user.is_owner:
        prop.num_views += 1
        prop.save()

    is_rented = Rented.objects.filter(prop_id=prop).exists()
    is_rented_by_me = False
    if curr_visitor:
        is_rented_by_me = Rented.objects.filter(prop_id=prop, visitor_id=curr_visitor).exists()

    reviews = Review.objects.filter(prop_id=prop).order_by('-id')

    if not current_user.is_owner and request.method == "POST":
        if 'rent' in request.POST and not is_rented:
            rented = Rented.objects.create(
                prop_id=prop,
                visitor_id=curr_visitor,
                owner_id=prop.owner,
                rent_to_be_paid=prop.price
            )
            rented.save()
            return redirect('dashboard')

        if 'form' in request.POST:
            comment = request.POST.get('comment', '').strip()
            rating_val = request.POST.get('rating', '5')
            try:
                rating = int(rating_val)
            except ValueError:
                rating = 5
            if comment:
                rev = Review.objects.create(
                    visitor_id=curr_visitor,
                    prop_id=prop,
                    comment=comment,
                    rating=rating
                )
                rev.save()
                all_revs = Review.objects.filter(prop_id=prop)
                prop.avg_rating = int(sum(r.rating for r in all_revs) / all_revs.count())
                prop.save()
                return redirect('prop', id=id)

    context = {
        'prop': prop,
        'review': reviews,
        'disp_comment': not current_user.is_owner,
        'is_rented': is_rented,
        'is_rented_by_me': is_rented_by_me,
        'error': error
    }
    return render(request, "view_property.html", context)

##########################################################

@login_required
def add_property(request):
    error = None
    current_user = request.user
    if not current_user.is_owner:
        return render(request, "addProperty.html", {'error': "Please login as a property owner to list properties."})

    try:
        current_owner = Owner.objects.get(user=current_user)
    except Owner.DoesNotExist:
        current_owner = Owner.objects.create(user=current_user, owner_name=current_user.username, num_properties=0)

    if request.method == "POST":
        title = request.POST.get('title', '').strip()
        property_type = request.POST.get('property_type', 'Apartment').strip()
        description = request.POST.get('description', '').strip()
        price_val = request.POST.get('price', '').strip()
        location = request.POST.get('location', '').strip()
        bedrooms = request.POST.get('bedrooms', '1').strip()
        bathrooms = request.POST.get('bathrooms', '1').strip()
        area_sqft = request.POST.get('area_sqft', '500').strip()
        contact_number = request.POST.get('contact_number', '').strip()
        furnishing = request.POST.get('furnishing', 'Semi-Furnished').strip()
        preferred_tenant = request.POST.get('preferred_tenant', 'Any').strip()
        parking = request.POST.get('parking', 'Two-Wheeler').strip()
        amenities = request.POST.get('amenities', 'Wi-Fi, 24x7 Water').strip()
        image = request.FILES.get('image')

        if not description or not price_val or not location:
            error = "Please fill in all required fields (Description, Price, Location)."
            return render(request, "addProperty.html", {'error': error})

        try:
            price = int(price_val)
            bed = int(bedrooms) if bedrooms else 1
            bath = int(bathrooms) if bathrooms else 1
            sqft = int(area_sqft) if area_sqft else 500
        except ValueError:
            error = "Price, Bedrooms, Bathrooms and Area must be valid numbers."
            return render(request, "addProperty.html", {'error': error})

        if not title:
            title = f"{property_type} in {location}"

        new_prop = Property.objects.create(
            owner=current_owner,
            title=title,
            property_type=property_type,
            description=description,
            price=price,
            location=location,
            bedrooms=bed,
            bathrooms=bath,
            area_sqft=sqft,
            contact_number=contact_number,
            furnishing=furnishing,
            preferred_tenant=preferred_tenant,
            parking=parking,
            amenities=amenities,
            image=image
        )
        new_prop.save()
        current_owner.num_properties = Property.objects.filter(owner=current_owner).count()
        current_owner.save()
        return redirect('ownerdashboard')

    return render(request, "addProperty.html", {'error': error})

##########################################################

def logout_user(request):
    logout(request)
    return redirect("home")

def home(request):
    featured_properties = Property.objects.all().order_by('-id')[:6]
    return render(request, "home.html", {'featured_properties': featured_properties})

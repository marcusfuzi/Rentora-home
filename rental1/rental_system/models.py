from django.db import models
from django.contrib.auth.models import AbstractUser
from django.conf import settings
# Create your models here.

class User(AbstractUser):
	is_owner = models.BooleanField(default = False)
	def __str__(self):
		res = str(self.username)
		return res

class Property(models.Model):
	id = models.AutoField(primary_key = True) 
	title = models.CharField(max_length = 120, default = '', blank = True)
	property_type = models.CharField(max_length = 50, default = 'Apartment', blank = True)
	description = models.TextField(default = '', blank = False)
	price = models.IntegerField(blank = False)
	location = models.CharField(max_length = 50, blank = False, null = False) 
	bedrooms = models.IntegerField(default = 2, blank = True)
	bathrooms = models.IntegerField(default = 1, blank = True)
	area_sqft = models.IntegerField(default = 1000, blank = True)
	contact_number = models.CharField(max_length = 20, default = '', blank = True)
	furnishing = models.CharField(max_length = 40, default = 'Semi-Furnished', blank = True)
	preferred_tenant = models.CharField(max_length = 40, default = 'Any', blank = True)
	parking = models.CharField(max_length = 40, default = 'Two-Wheeler', blank = True)
	amenities = models.CharField(max_length = 200, default = 'Wi-Fi, 24x7 Water', blank = True)
	image = models.ImageField(upload_to = 'properties/', null = True, blank = True)
	num_views = models.IntegerField(default=0)
	avg_rating = models.IntegerField(default=0)
	owner = models.ForeignKey('rental_system.Owner', on_delete=models.CASCADE, null = True)

	def __str__(self):
		return f"{self.title or 'Property #' + str(self.id)} ({self.location})"

	@property
	def display_image_url(self):
		if self.image:
			try:
				return self.image.url
			except Exception:
				pass
		return '/static/images/house1.jpg'

class Visitor(models.Model):
	user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
	id = models.AutoField(primary_key = True)
	profile = models.TextField(blank = False)
	pref_location = models.CharField(max_length = 30) 
	def __str__(self):
		return self.user.username;

class Owner(models.Model):
	user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
	id = models.AutoField(primary_key = True)
	num_properties = models.IntegerField(default=0)
	owner_name = models.CharField(max_length=30)
	def __str__(self):
		return self.owner_name

class Rented(models.Model):
	owner_id = models.ForeignKey('rental_system.Owner',on_delete = models.CASCADE )
	prop_id = models.ForeignKey('rental_system.Property',on_delete = models.CASCADE )
	visitor_id = models.ForeignKey('rental_system.Visitor',on_delete = models.CASCADE )
	rent_to_be_paid = models.IntegerField(default=0)
	def __str__(self):
		return "%s %s %s" % (self.owner_id, self.visitor_id, self.rent_to_be_paid)

class Review(models.Model):
	prop_id = models.ForeignKey('rental_system.Property',on_delete = models.CASCADE )
	visitor_id = models.ForeignKey('rental_system.Visitor',on_delete = models.CASCADE )
	rating = models.IntegerField(default=0)
	comment = models.CharField(max_length = 100,blank = False)
	def __str__(self):
		return "%s %s %s" % (self.comment, self.rating, self.visitor_id)

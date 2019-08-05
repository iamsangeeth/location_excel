from django.shortcuts import render, HttpResponse
from .forms import DocumentForm
import os
from django.conf import settings
import openpyxl
import googlemaps


# Create your views here.
def home(request):
	file_form = DocumentForm()
	context = {
		"title":"Hello World!",
		"content":" Welcome to the homepage.",
		"form": file_form,
	}
	return render(request, "home.html", context)

def process(request):
	# Handle file upload
	if request.method == 'POST':
		form = DocumentForm(request.POST, request.FILES)
		if form.is_valid():
			excel_file = request.FILES['docfile']
			wb = openpyxl.load_workbook(excel_file)
			file_dir = os.path.join(settings.BASE_DIR, 'resource')
			worksheet = wb["Sheet0"]
			excel_data = list()
			gmaps = googlemaps.Client(key='AIzaSyCMUJJy-98yTy2Y7P5tarHrUDex2LnoDFk')
			for row in worksheet.iter_rows():
				if row[0]:
					address = ''
					for i in range(0,len(row)):
						if row[i].value:
							if address:
								address = address + ', '+ str(row[i].value)
							else:
								address = str(row[i].value)
						else:
							break
					print(address)
					if address:
						geocode_result = gmaps.geocode(address)
						row[i].value = geocode_result[0]['geometry']['location']['lat']
						row[i+1].value = geocode_result[0]['geometry']['location']['lng']

			out_file_path = os.path.join(file_dir,"outfile.xlsx")
			wb.save(out_file_path)
			# file_path = os.path.join(settings.MEDIA_ROOT, path)
			if os.path.exists(out_file_path):
				with open(out_file_path, 'rb') as fh:
					response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
					response['Content-Disposition'] = 'inline; filename=' + os.path.basename(out_file_path)
					return response
 
	else:
		form = DocumentForm()
	return render(request, 'home.html')
#import dependencies
import os	
import shutil								
import openpyxl		
import datetime												
from PIL import Image													
from PIL import ImageFont
from PIL import ImageDraw
import argparse

#check whether the cli argument is a valid file or not
def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        return arg

#parse the arguments
parser = argparse.ArgumentParser(description='A script to generate member certificates for clubs.')
parser.add_argument('-f', '--file', required=True,help='input file with member details', metavar='',type=lambda x: is_valid_file(parser, x))
parser.add_argument('-i', '--image',required=False, default='volunteer.jpg',help='club member certificate template', metavar='',type=lambda x: is_valid_file(parser, x))
args = vars(parser.parse_args())

#get variable names
FILE = args['file']
IMAGE = args['image']

#set up the environment i.e. take care of folders and files to create
f = open(FILE,"r")
dir_name=f.readline().split(',')[2]+" Certificates Generated Here"

if not os.path.exists(dir_name):
	os.makedirs(dir_name)

#goto start of file
f.seek(0)

#iterate over each line in file and generate the certificates
for line in f:
	name,role,club,start_year,end_year,date,capacity=line.split(',')
	print(name)
	img=Image.open(IMAGE)	
	font=ImageFont.truetype("Crimson-Roman.ttf",size=60)
	
	W, H = img.size
	
	draw = ImageDraw.Draw(img)

	w, h = draw.textsize(name.title(), font=font)
	draw.text(((W-w)/2,1000),name.title(),(0,0,0),font=font)

	w, h = draw.textsize(role.title(), font=font)
	draw.text(((W-w)/2,1200),role.title(),(0,0,0),font=font)

	w, h = draw.textsize(club, font=font)
	draw.text(((W-w)/2,1340),club,(0,0,0),font=font)

	span = start_year+'-'+end_year
	w, h = draw.textsize(span, font=font)
	draw.text(((W-w)/2,1480),span,(0,0,0),font=font)

	w, h = draw.textsize(date, font=font)
	draw.text(((W-w)/2-750,1780),date,(0,0,0),font=font)

	w, h = draw.textsize(capacity, font=font)
	draw.text(((W-w)/2+750,1750),capacity,(0,0,0),font=font)

	img.save(dir_name+'/'+name+'.jpg')

#close the file
f.close()

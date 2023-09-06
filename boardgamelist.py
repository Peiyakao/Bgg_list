# -*- coding: utf-8 -*
import pygsheets
import traceback
import requests
from bs4 import BeautifulSoup
import json
import sys

SHT_URL = 'https://docs.google.com/spreadsheets/d/1ANthTc94WtMIXYHunXaLN3svWYZg6CwAIyIH8VJVQD4/edit#gid=0'
SHT_JSON = 'boardgamelist.json'
SHT_NAME = 'bgg_list'


def update_rating(cell, rating):
	try:
		rating = round(float(rating), 1)
		if rating:
			cell.value = rating
			if rating >= 5 and rating < 6:
				cell.color = (0.8509804, 0.8235294, 0.9137255, 0)
			elif rating >= 6 and rating < 7:
				cell.color = (0.7882353, 0.85490197, 0.972549, 0)
			elif rating >= 7 and rating < 8:
				cell.color = (0.43529412, 0.65882355, 0.8627451, 0)
			elif rating >= 8 and rating < 9:
				cell.color = (0.5764706, 0.76862746, 0.49019608, 0)
			elif rating >= 9:
				cell.color = (0.41568628, 0.65882355, 0.30980393, 0)
			else:
				cell.color = (1, 1, 1, 0)
		else:
			cell.value = ''
			cell.color = (1, 1, 1, 0)
	except:
		print traceback.format_exc()

def update_weight(cell, weight):
	try:
		weight = round(float(weight), 2)
		if weight:
			cell.value = weight
			if weight >= 1 and weight < 2:
				cell.color = (0.8509804, 0.91764706, 0.827451, 0)
			elif weight >= 2 and weight < 3:
				cell.color = (0.5764706, 0.76862746, 0.49019608, 0)
			elif weight >= 3 and weight < 4:
				cell.color = (0.9764706, 0.79607844, 0.6117647, 0)
			elif weight >= 4:
				cell.color = (0.9019608, 0.5686275, 0.21960784, 0)
			else:
				cell.color = (1, 1, 1, 0)
		else:
			cell.value = ''
			cell.color = (1, 1, 1, 0)
	except:
		print traceback.format_exc()

def update_ranking(cell, ranking):
	try:
		ranking = int(ranking)
		if ranking:
			cell.value = int(ranking)
		else:
			cell.value = ''
		
		cell.color = (1, 1, 1, 0)
	except:
		print traceback.format_exc()

def update_string(cell, s):
	try:
		cell.value = s
		cell.color = (1, 1, 1, 0)
	except:
		print traceback.format_exc()

def html_parser(url):
	paras = {}
	script_var_dict = {}
	response = requests.get(url)
	soup = BeautifulSoup(response.text, 'html.parser')
	#print soup.prettify()
	script = soup.find('script', text=lambda text: text and 'GEEK' in text)
	script_var_list = script.string.split('\n')
	
	# find GEEK.geekitemPreload in script var
	for script_var in script_var_list:
		script_var = script_var.strip()
		if script_var.startswith('GEEK.geekitemPreload'):
			script_var_dict = json.loads(script_var[23:-1])
			break
	
	# rank
	for rankinfo in script_var_dict['item']['rankinfo']:
		if rankinfo['shortprettyname'] == 'Overall Rank':
			paras['rank'] = rankinfo['rank']
			break
	
	# name
	paras['name'] = script_var_dict['item']['name']
	
	# players
	paras['players'] = '%s-%s' % (script_var_dict['item']['minplayers'], script_var_dict['item']['maxplayers'])
	
	# stats: average, avgweight
	paras['average'] = script_var_dict['item']['stats']['average']
	paras['avgweight'] = script_var_dict['item']['stats']['avgweight']
	
	# category
	category_string = ''
	for category in script_var_dict['item']['links']['boardgamecategory']:
		if category_string:
			# print type(category_string)
			category_string += unicode('、', 'utf-8')
		
		category_string += category['name']
	
	paras['category_string'] = category_string
	
	# subdomain
	subdomain_string = ''
	for subdomain in script_var_dict['item']['links']['boardgamesubdomain']:
		if subdomain_string:
			subdomain_string += unicode('、', 'utf-8')
		
		subdomain_string += subdomain['name'].replace(' Games', '')
	
	paras['subdomain'] = subdomain_string
	
	# playtime
	min = script_var_dict['item']['minplaytime']
	max = script_var_dict['item']['maxplaytime']
	if min == max:
		paras['playtime'] = min
	else:
		paras['playtime'] = min + '-' + max
	
	return paras

def main():
	def update_row(wks_web, i):
		bggid_idx = 'A%d' % i
		bggid_cell = wks_web.cell(bggid_idx)
		if bggid_cell.value:
			paras = html_parser('https://boardgamegeek.com/boardgame/' + bggid_cell.value)
			print i, paras
			
			# update web display wks
			if 'name' in paras:
				idx = 'C%d' % i
				w_cell = wks_web.cell(idx)
				update_string(w_cell, paras['name'])
			
			if 'rank' in paras:
				idx = 'D%d' % i
				w_cell = wks_web.cell(idx)
				update_ranking(w_cell, paras['rank'])
			
			if 'average' in paras:
				idx = 'E%d' % i
				w_cell = wks_web.cell(idx)
				update_rating(w_cell, paras['average'])
			
			if 'avgweight' in paras:
				idx = 'F%d' % i
				w_cell = wks_web.cell(idx)
				update_weight(w_cell, paras['avgweight'])
			
			if 'players' in paras:
				idx = 'G%d' % i
				w_cell = wks_web.cell(idx)
				update_string(w_cell, paras['players'])
			
			if 'playtime' in paras:
				idx = 'H%d' % i
				w_cell = wks_web.cell(idx)
				update_string(w_cell, paras['playtime'])
			
			if 'subdomain' in paras:
				idx = 'I%d' % i
				w_cell = wks_web.cell(idx)
				update_string(w_cell, paras['subdomain'])
	
	# update account info & excel info here
	gc = pygsheets.authorize(service_file=SHT_JSON, seconds_per_quota=5)
	sht = gc.open_by_url(SHT_URL)
	wks_list = sht.worksheets()
	wks_web = sht.worksheet_by_title(SHT_NAME)
	
	# print sys.argv
	# print eval(sys.argv[1])
	# print type(eval(sys.argv[1]))
	# paras = json.loads(sys.argv[1])
	paras = eval(sys.argv[1])
	if 'update' in paras:
		if 'all' in paras['update']:
			update_list = range(2, wks_web.rows+1)
		else:
			update_list = paras['update']
		
		if '..' in update_list:
			x,y = update_list.split('..')
			update_list = range(int(x), int(y)+1)
		
		for i in update_list:
			update_row(wks_web, int(i))

"""
 example:
	python boardgamelist.py "{'update': ['all']}"
	python boardgamelist.py "{'update': [2]}"
	python boardgamelist.py "{'update': [1,3,5]}"
	python boardgamelist.py "{'update': ['6..8']}"
"""
if __name__ == '__main__':
	main()
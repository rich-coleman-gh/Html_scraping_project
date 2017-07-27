# coding: utf-8
from lxml import html
import requests

page = requests.get('https://www.sec.gov/Archives/edgar/data/1442145/000144214517000009/ex991q42016.htm')
tree = html.fromstring(page.content)


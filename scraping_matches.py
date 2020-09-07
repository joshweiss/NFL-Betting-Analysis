from bs4 import BeautifulSoup
import requests
from sportsreference.nfl.teams import Teams as nflTeams
import random


source = requests.get('http://www.footballlocks.com/nfl_point_spreads.shtml').text

soup = BeautifulSoup(source, 'lxml')

teams = nflTeams()

table = soup.find('table', cols = "4")


matches = table.find_all('tr')

"""mon_night = soup.find_all('table', cols = "4")[1].find_all('tr')[1]
matches.append(mon_night)"""


for match in matches[1:]:
    items = match.find_all('td')
    fav = items[1].text.split(' ')[-1]
    if 'PK' in items[2].text:
        spread = 0.0
    else:
        spread = abs(float(items[2].text))
    und = items[3].text.split(' ')[-1]
    if fav == "Bay":
        fav = items[1].text.split(' ')[-2] + ' ' + items[1].text.split(' ')[-1]
    if und == "Bay":
        und = items[3].text.split(' ')[-2] + ' ' + items[3].text.split(' ')[-1]

    found1, found2 = False, False
    for team in teams:
        if fav in team.name:
            fav = team.name
            found1 = True
        if und in team.name:
            und = team.name
            found2 = True
        if found1 and found2:
            break

    if items[1].text.split(' ')[0] == 'At':
        home = fav
        away = und
    else:
        home = und
        away = fav

    if spread.is_integer():
        spread += random.choice([-0.5,0.5])


    print(away + ' @ ' + home + ' (' + fav.split(' ')[-1] + ' -' + str(spread) + ')')

    print('------------------------------------')

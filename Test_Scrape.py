from sportsreference.nfl.teams import Teams as nflTeams
from sportsreference.nfl.boxscore import Boxscores


scores = Boxscores(7, 2018).games['{b}-{c}'.format(b=7, c=2018)]

for score in scores:
    print (score)
def bye_data():
    yearly_bye_wins = [0 for i in range(13)]
    yearly_bye_losses = [0 for i in range(13)]

    for team in team_results:
        for y,year in enumerate(team_results[team]):
            for i,week in enumerate(year):
                if i > 1:
                    if year[i-2] == year[i-1]:
                        if week > year[i-1]:
                            yearly_bye_wins[y] += 1
                        else:
                            yearly_bye_losses[y] += 1

    for col in range(4,17):
        sh_overall.cell(12, col).value = yearly_bye_wins[col-4]
        sh_overall.cell(13, col).value = yearly_bye_losses[col-4]

def fill_data_sheet():
    row = 2
    for team in team_results:
        for sheet in wk_analysis:
            if sheet.title != 'Overall':
                for col in range(2, 19):
                    sheet.cell(row, col).value = team_results[team][2018 - int(sheet.title)][col - 1]
                row += 1
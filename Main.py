import pandas as pd
import matplotlib.pylab as plt


file_input = 'data_to_analyze/vgsales.csv'
df = pd.read_csv(file_input)

output_file = 'result/analyzed.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Sheets Names
    df.to_excel(writer, sheet_name='RAW')
    pd.DataFrame().to_excel(writer, sheet_name='DA')

    # Frameworks
    workbook = writer.book
    worksheet = writer.sheets['DA']

    # Pandas Formulas
    # Fix Cells
    cells = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})
    worksheet.conditional_format(
        'A1:G100', {'type': 'no_blanks', 'format': cells})

    # Total Sales By Platform
    analysis = [{'group_col': 'Platform', 'col': [
        'NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'math': 'sum', 'by': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'row': 2, 'title': 'Sales in millions of copies sold by platform (Top 15)'},
        {'group_col': 'Genre', 'col': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'math': 'sum', 'by': [
            'NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'row': 21, 'title': 'Sales in millions of copies sold by Genre'},
        {'group_col': 'Publisher', 'col': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'math': 'sum', 'by': [
            'NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'row': 37, 'title': 'Sales in millions of copies sold by publisher (Top 15)'},
        {'group_col': 'Year', 'col': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'math': 'sum', 'by': [
            'NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'row': 56, 'title': 'Sales in millions of copies sold by year'},
        {'group_col': 'Name', 'col': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'math': 'sum', 'by': ['NA_Sales', 'EU_Sales', 'JP_Sales', 'Other_Sales'], 'row': 75, 'title': 'Sales in millions of copies sold by name (Top 15)'}]

    for item in analysis:
        if item['math'] == 'sum':
            result = df.groupby(item['group_col'])[
                item['col']].sum().sort_values(by=item['by'], ascending=False).head(15)

        result.to_excel(writer, sheet_name='DA',
                        startrow=item['row'], startcol=0)
        worksheet.merge_range(f"A{item['row']}:E{item['row']}", item['title'])

    # Publisher With Most Releases Top 10
    most_releases = df['Publisher'].value_counts(
    ).sort_values(ascending=False).head(10)

    plt.pie(most_releases,
            labels=most_releases.index, autopct='%1.1f%%', colors=['gold', 'blue', 'red', 'yellow', 'purple', 'forestgreen', 'tomato', 'silver', 'orange', 'brown'])
    plt.title('Top 10 publishers with most releases')
    plt.savefig('pie.png')
    plt.close()
    worksheet.insert_image('H3', 'pie.png')

    # Global Sales Through The Years By Genre
    global_sales = [{'group_cols': ['Year', 'Genre'], 'target_cols': 'Global_Sales',
                     'math': 'sum', 'titles': 'Global sales over time by genre top 10', 'xlabel': 'Year', 'ylabel': 'Global sales in millions of copies', 'image': 'plot1.png', 'rows': 28},
                    {'group_cols': ['Year', 'Publisher'], 'target_cols': 'Global_Sales', 'math': 'sum', 'titles': 'Global sales over time by publisher top 10', 'xlabel': 'Year', 'ylabel': 'Global sales in millions of copies', 'image': 'line2.png',
                     'rows': 56},
                    {'group_cols': ['Year', 'Platform'], 'target_cols': 'Global_Sales', 'math': 'sum', 'titles': 'Global sales over time by platform top 10',
                     'xlabel': 'Year', 'ylabel': 'Global sales in millions of copies', 'image': 'line3.png', 'rows': 82}]

    for items in global_sales:
        if items['math'] == 'sum':
            output_name = df.groupby(items['group_cols'])[
                items['target_cols']].sum().unstack()
            top10_columns = output_name.sum().nlargest(10).index
            full_output = output_name[top10_columns]

            full_output.plot(kind='line', figsize=(10, 5))
            plt.title(items['titles'])
            plt.xlabel(items['xlabel'])
            plt.ylabel(items['ylabel'])
            plt.savefig(items['image'])
            plt.close()
            worksheet.insert_image(f"H{items['rows']}", items['image'])

    # Top 10 Games With Most Global Sales
    games_global = df.groupby(
        'Name')['Global_Sales'].sum().sort_values(ascending=False).head(10)

    games_global.plot(kind='barh', figsize=(8, 5))
    plt.title('Top 10 games by global sales')
    plt.xlabel('Global sales in millions')
    plt.ylabel('Game title')
    plt.subplots_adjust(left=0.40)
    plt.savefig('bar.png')
    plt.close()
    worksheet.insert_image('T2', 'bar.png')

    # Dashboard Name
    dst = workbook.add_format({'bold': True, 'align': 'center',
                               'font_size': 16, 'bg_color': "#0396f8"})
    worksheet.merge_range('A1:Z1', 'Global Video Game Sales Analytics', dst)

    worksheet.hide_gridlines(2)
    worksheet.autofit()

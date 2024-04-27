import xlsxwriter


def export_check(text):
    text = [[(k.split('\t')[0], int(k.split('\t')[1]), int(k.split('\t')[2]))
             for k in t.split('\n') if k] for t in text.split('---')]
    workbook = xlsxwriter.Workbook('res.xlsx')
    for i in text:
        worksheet = workbook.add_worksheet()
        data = []
        for row, (name, price, quality) in enumerate(i):
            flag = False
            for j in range(len(data)):
                if data[j][0] == name and data[j][1] == price:
                    data[j][2] += quality
                    flag = True
            if not flag:
                data.append([name, price, quality])
        data = sorted(data, key=lambda x: x[0])
        for j in range(len(data)):
            worksheet.write(j, 0, data[j][0])
            worksheet.write(j, 1, data[j][1])
            worksheet.write(j, 2, data[j][2])
            worksheet.write(j, 3, f'=B{j + 1} * C{j + 1}')
        worksheet.write(len(set(i)), 0, 'Итого')
        worksheet.write(len(set(i)), 3, f'=SUM(D1:D{len(set(i))})')
    workbook.close()



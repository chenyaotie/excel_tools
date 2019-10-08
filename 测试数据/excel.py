#encoding=utf-8
from xlwt import Workbook, easyxf

def show_color(sheet):
        colNum = 6
        width = 5000
        height = 500
        colors = ['aqua','black','blue','blue_gray','bright_green','brown','coral','cyan_ega','dark_blue','dark_blue_ega','dark_green','dark_green_ega','dark_purple','dark_red',
                'dark_red_ega','dark_teal','dark_yellow','gold','gray_ega','gray25','gray40','gray50','gray80','green','ice_blue','indigo','ivory','lavender',
                'light_blue','light_green','light_orange','light_turquoise','light_yellow','lime','magenta_ega','ocean_blue','olive_ega','olive_green','orange','pale_blue','periwinkle','pink',
                'plum','purple_ega','red','rose','sea_green','silver_ega','sky_blue','tan','teal','teal_ega','turquoise','violet','white','yellow']

        for colorIndex in range(len(colors)):
                rowIndex = colorIndex / colNum
                colIndex = colorIndex - rowIndex*colNum
                sheet.col(colIndex).width = width
                sheet.row(rowIndex).set_style(easyxf('font:height %s;'%height)) 
                color = colors[colorIndex]
                whiteStyle = easyxf('pattern:pattern solid, fore_colour %s;'
                                        'align: vertical center, horizontal center;'
                                        'font: bold true, colour white;' % color)
                blackStyle = easyxf('pattern:pattern solid, fore_colour %s;'
                                        'align: vertical center, horizontal center;'
                                        'font: bold true, colour black;' % color)


                if color == 'black':
                        sheet.write(rowIndex, colIndex, color, style = whiteStyle)
                else:
                        sheet.write(rowIndex, colIndex, color, style = blackStyle)

def show_size(sheet):
        widthStart = 100
        widthInterval = 100
        colNum = 255
        heightStart = 100
        heightInterval = 5
        rowNum = 255
        styles = (easyxf('pattern:pattern solid, fore_colour gray50;'
                        'align: vertical center, horizontal center;'
                        'font: bold true, colour white;'),
                easyxf('pattern:pattern solid, fore_colour gray80;'
                        'align: vertical center, horizontal center;'
                        'font: bold true, colour white;'))
        for rowIndex in range(rowNum):
                height = heightStart + heightInterval*rowIndex
                sheet.row(rowIndex).set_style(easyxf('font:height %s;'%height))
                styleIndex = rowIndex%2
                for colIndex in range(colNum):
                        width = widthStart + widthInterval*colIndex
                        sheet.col(colIndex).width = width
                        sheet.write(rowIndex, colIndex, '%sx%s'%(width,height), style = styles[styleIndex])
                        styleIndex = int(not styleIndex)


if __name__ == '__main__':
        book = Workbook(encoding='utf-8')
        colorSheet = book.add_sheet('colors')
        sizeSheet = book.add_sheet('size')
        show_color(colorSheet)
        show_size(sizeSheet)
        styleFile = 'excel_styles.xls'
        book.save(styleFile)
        print 'saved to "%s"' % styleFile

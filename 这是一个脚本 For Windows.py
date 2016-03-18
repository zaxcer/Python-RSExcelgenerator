#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlwt
import os

class RSExcel(object):

	def __init__(self, title = 'default', number = 6, path = '/Users/Zhangchi/Desktop/Test.xlsx'):

		self.title = title
		self.number = number
		self.path = path
		self.file = xlwt.Workbook()

	def write(self):
		sheet = self.file.add_sheet('sheet%1', cell_overwrite_ok=True)
		###调整最中间的列宽，以适应纸张大小
		sheet.col(2).width = 500
		#按顺序按区块写入信息
		for n in range(int(self.number)):

			if n%2 == 0:
				coor_x = n * 2
				coor_y = 0
			elif n%2 == 1:
				coor_x = (n - 1) * 2
				coor_y = 3
			else:
				coor_x = 0
				coor_y = 0
				print('coorX & coorY have some problems')
			
			self.write_to_block(sheet, coor_x, coor_y, n)

		self.file.save(self.path)

	def write_to_block(self, sheet, coordinate_x = 0, coordinate_y = 0, block_index = 0):
		#Cell长宽设置时乘以该系数，调整合适大小
		factor = 50

		style1 = self.set_style('Arial Black', 8, 0, False)
		#SimSun
		style2 = self.set_style('SimSun', 48, 0, True)
		#'P NO.'的horz设为右对齐
		style3 = self.set_style('SimSun', 24, 0, True, horz=xlwt.Alignment.HORZ_RIGHT)
		style4 = self.set_style('SimSun', 24, 0, True)

		sheet.write(coordinate_x+1, coordinate_y, 'RHODE&SCHWARZ', style1)
		sheet.write(coordinate_x+2, coordinate_y, 'BOX', style2)
		sheet.write(coordinate_x+2, coordinate_y+1, str(block_index+1)+'/'+str(self.number), style2)
		sheet.write(coordinate_x+3, coordinate_y, 'P NO.', style3)
		sheet.write(coordinate_x+3, coordinate_y+1, self.title, style4)
		#set sheet width
		sheet.col(coordinate_y).width = 100 * factor
		sheet.col(coordinate_y+1).width = 150 * factor
		#set sheet height
		for x in range(coordinate_x, coordinate_x+4):
			#高度的调整需要先把height._mismatch属性设为True
			sheet.row(x).height_mismatch = True
		sheet.row(coordinate_x).height = 4 * factor
		sheet.row(coordinate_x+1).height = 25 * factor
		sheet.row(coordinate_x+2).height = 50 * factor
		sheet.row(coordinate_x+3).height = 25 * factor

	def set_style(self, font_name, font_height, font_color, font_bold, horz = xlwt.Alignment.HORZ_CENTER):
		style = xlwt.XFStyle()

		#set font
		font = xlwt.Font()
		font.name = font_name
		font.height = font_height*20
		font.colour_index = font_color
		font.bold = font_bold

		#set alignment
		alignment = xlwt.Alignment()
		alignment.horz = horz  #horizontal 参数默认水平居中
		alignment.vert = xlwt.Alignment.VERT_CENTER
		alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT

		style.font = font
		style.alignment = alignment

		return style


if __name__ == '__main__':

	ModelTitle = input('请输入型号: ')
	ModelNumber = input('请输入打印数量: ')

	path = 'D:\RSlabel.xls'

	RSexcel = RSExcel(ModelTitle, ModelNumber, path)
	RSexcel.write()

	os.system('D:\RSlabel.xls')
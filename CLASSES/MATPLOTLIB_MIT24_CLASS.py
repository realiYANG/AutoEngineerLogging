# coding=utf-8
import io

import las
import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlrd
import xlwt
from matplotlib.widgets import MultiCursor, RadioButtons, SpanSelector
from PyQt5.QtWidgets import (QApplication, QFileDialog, QMainWindow,
                             QTableWidgetItem)


class MATPLOTLIB_MIT24:
    def __init__(self, fileDir):
        self.damage_Tag = ''
        self.lines = []  # 生成的成果list
        self.log = las.LASReader(fileDir, null_subs=np.nan)
        self.fig1 = plt.figure('MIT油套管快速评价系统', figsize=(12, 8))
        xls = xlrd.open_workbook(".\\casing_data.xls")
        table = xls.sheet_by_name('Sheet1')
        # 注意下面几个的类型为excel单元格对象
        self.outer_diameter = table.cell(0, 2)
        self.inner_diameter = table.cell(0, 3)
        self.thickness = table.cell(0, 4)

        self.scale_left = float(self.inner_diameter.value) / 2 - 20
        self.scale_right = float(self.inner_diameter.value) / 2 + 80

        self.scale_left_min = float(self.inner_diameter.value) - 30
        self.scale_right_max = float(self.inner_diameter.value) + 30

        # 定义RadioButtons
        axcolor = 'lightgoldenrodyellow'
        rax = plt.axes([0.75, 0.05, 0.12, 0.09], facecolor=axcolor)
        radio = RadioButtons(rax, (u'Penetration', u'Projection', u'Transformation', u'Holes-detection', u'Type-suggestion'), active=-1, activecolor='red')
        plt.subplots_adjust(bottom=0.15, top=0.95, right=0.9, left=0.10, wspace=0.60)
        radio.on_clicked(self.actionfunc)
        #####################################################################################
        # 坐标轴1
        self.ax1 = plt.subplot(141)
        # 下面赋值加逗号是为了使得type(self.line1)为matplotlib.lines.Line2D对象，而不是list
        self.line1, = self.ax1.plot(self.log.data['FING01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['FING06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['FING11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['FING16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax1.plot(self.log.data['FING21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax1.plot(self.log.data['FING24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax1.set_xlim(self.scale_left, self.scale_right)
        self.ax1.set_ylim(self.log.start, self.log.stop)
        self.ax1.invert_yaxis()

        span1 = SpanSelector(self.ax1, self.onselect1, 'vertical', useblit=False,
                             rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)
        # plt.ylabel(self.log.curves.DEPT.descr + " (%s)" % self.log.curves.DEPT.units)
        # plt.xlabel(self.log.curves.FING01.descr + " (%s)" % self.log.curves.FING01.units)
        # plt.title(self.log.well.WELL.data)
        plt.ylabel('Measured Depth(m)')
        plt.title('Original')

        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        plt.grid()
        #####################################################################################
        # 坐标轴2
        self.ax2 = plt.subplot(142)
        self.line2, = self.ax2.plot(self.log.data['FING01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['FING06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['FING11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['FING16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax2.plot(self.log.data['FING21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax2.plot(self.log.data['FING24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax2.set_xlim(self.scale_left, self.scale_right)
        self.ax2.set_ylim(self.log.start, self.log.stop)
        self.ax2.invert_yaxis()

        span2 = SpanSelector(self.ax2, self.onselect2, 'vertical', useblit=False,
                             rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)
        plt.title('Middle')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax2.grid()
        #####################################################################################
        # 坐标轴3
        self.ax3 = plt.subplot(143)
        self.line3, = self.ax3.plot(self.log.data['FING01'], self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING02'] + 2.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING03'] + 5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING04'] + 7.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING05'] + 10, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['FING06'] + 12.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING07'] + 15, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING08'] + 17.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING09'] + 20, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING10'] + 22.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['FING11'] + 25, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING12'] + 27.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING13'] + 30, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING14'] + 32.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING15'] + 35, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['FING16'] + 37.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING17'] + 40, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING18'] + 42.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING19'] + 45, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING20'] + 47.5, self.log.data['DEPT'], 'r-', lw=0.3)
        self.ax3.plot(self.log.data['FING21'] + 50, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING22'] + 52.5, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING23'] + 55, self.log.data['DEPT'], 'g-', lw=0.3)
        self.ax3.plot(self.log.data['FING24'] + 57.5, self.log.data['DEPT'], 'g-', lw=0.3)

        self.ax3.set_xlim(self.scale_left, self.scale_right)
        self.ax3.set_ylim(self.log.start, self.log.stop)
        self.ax3.invert_yaxis()

        # self.span3_cyan = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
        #                           rectprops=dict(alpha=0.5, facecolor='cyan'), span_stays=True)

        plt.title('Large')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax3.grid()
        #####################################################################################
        # 坐标轴4
        self.ax4 = plt.subplot(144)
        self.ax4.plot(self.log.data['MAXDIA'], self.log.data['DEPT'], 'r--', lw=0.3)
        self.ax4.plot(self.log.data['MINDIA'], self.log.data['DEPT'], 'b-', lw=0.3)
        self.ax4.plot(self.log.data['AVEDIA'], self.log.data['DEPT'], 'k--', lw=0.3)

        self.ax4.set_xlim(self.scale_left_min - 20, self.scale_right_max + 20)
        self.ax4.set_ylim(self.log.start, self.log.stop)
        self.ax4.invert_yaxis()

        plt.title('Min-Max')
        plt.gca().spines['bottom'].set_position(('data', 0))
        plt.gca().spines['top'].set_position(('data', 0))
        self.ax4.grid()

        multi = MultiCursor(plt.gcf().canvas, (self.ax1, self.ax2, self.ax3, self.ax4), color='r', lw=1,
                            horizOn=True, vertOn=False)
        #########################

        plt.show()

    def onselect1(self, ymin, ymax):
        indmin, indmax = np.searchsorted(self.log.data['DEPT'], (ymin, ymax))
        indmax = min(len(self.log.data['DEPT']) - 1, indmax)
        thisx = self.log.data['FING01'][indmin:indmax]
        thisy = self.log.data['DEPT'][indmin:indmax]
        self.line2.set_data(thisx, thisy)
        self.ax2.set_ylim(thisy[-1], thisy[0])
        self.line3.set_data(thisx, thisy)
        self.ax3.set_ylim(thisy[-1], thisy[0])
        self.ax4.set_ylim(thisy[-1], thisy[0])
        plt.gcf().canvas.draw()

    def onselect2(self, ymin, ymax):
        indmin, indmax = np.searchsorted(self.log.data['DEPT'], (ymin, ymax))
        indmax = min(len(self.log.data['DEPT']) - 1, indmax)
        thisx = self.log.data['FING01'][indmin:indmax]
        thisy = self.log.data['DEPT'][indmin:indmax]
        self.line3.set_data(thisx, thisy)
        self.ax3.set_ylim(thisy[-1], thisy[0])
        self.ax4.set_ylim(thisy[-1], thisy[0])
        plt.gcf().canvas.draw()

    def onselect3(self, ymin, ymax):
        ymin = round(ymin, 2)
        ymax = round(ymax, 2)
        if ymin != ymax:
            self.lines = []  # 清空一下
            print('井段为：', ymin, '-', ymax)
            self.lines.append(ymin)
            self.lines.append(ymax)
        elif ymin == ymax:
            print('极值深度为：', ymax)
            ymax = round(ymax, 2)
            self.lines.append(ymax)
            index = np.where(abs(self.log.data['DEPT'] - ymax) <= float(abs(self.log.step) - 0.001))
            index = index[0][0]
            FING01_value = self.log.data['FING01'][index]
            FING02_value = self.log.data['FING02'][index]
            FING03_value = self.log.data['FING03'][index]
            FING04_value = self.log.data['FING04'][index]
            FING05_value = self.log.data['FING05'][index]
            FING06_value = self.log.data['FING06'][index]
            FING07_value = self.log.data['FING07'][index]
            FING08_value = self.log.data['FING08'][index]
            FING09_value = self.log.data['FING09'][index]
            FING10_value = self.log.data['FING10'][index]
            FING11_value = self.log.data['FING11'][index]
            FING12_value = self.log.data['FING12'][index]
            FING13_value = self.log.data['FING13'][index]
            FING14_value = self.log.data['FING14'][index]
            FING15_value = self.log.data['FING15'][index]
            FING16_value = self.log.data['FING16'][index]
            FING17_value = self.log.data['FING17'][index]
            FING18_value = self.log.data['FING18'][index]
            FING19_value = self.log.data['FING19'][index]
            FING20_value = self.log.data['FING20'][index]
            FING21_value = self.log.data['FING21'][index]
            FING22_value = self.log.data['FING22'][index]
            FING23_value = self.log.data['FING23'][index]
            FING24_value = self.log.data['FING24'][index]
            self.Min_value = self.log.data['MINDIA'][index]
            self.Ave_value = self.log.data['AVEDIA'][index]
            self.Max_value = self.log.data['MAXDIA'][index]
            print(str(self.Min_value), ' ', str(self.Ave_value), ' ', str(self.Max_value))
            if self.damage_Tag == 'Transformation':
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Transformation.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()
            elif self.damage_Tag == 'Penetration':
                All_Fingers = [FING01_value, FING02_value, FING03_value, FING04_value, FING05_value, FING06_value,
                               FING07_value,
                               FING08_value, FING09_value, FING10_value, FING11_value, FING12_value, FING13_value,
                               FING14_value,
                               FING15_value, FING16_value, FING17_value, FING18_value, FING19_value, FING20_value,
                               FING21_value,
                               FING22_value, FING23_value, FING24_value]
                All_Fingers_Dict = {'FING01_value': FING01_value,
                                    'FING02_value': FING02_value,
                                    'FING03_value': FING03_value,
                                    'FING04_value': FING04_value,
                                    'FING05_value': FING05_value,
                                    'FING06_value': FING06_value,
                                    'FING07_value': FING07_value,
                                    'FING08_value': FING08_value,
                                    'FING09_value': FING09_value,
                                    'FING10_value': FING10_value,
                                    'FING11_value': FING11_value,
                                    'FING12_value': FING12_value,
                                    'FING13_value': FING13_value,
                                    'FING14_value': FING14_value,
                                    'FING15_value': FING15_value,
                                    'FING16_value': FING16_value,
                                    'FING17_value': FING17_value,
                                    'FING18_value': FING18_value,
                                    'FING19_value': FING19_value,
                                    'FING20_value': FING20_value,
                                    'FING21_value': FING21_value,
                                    'FING22_value': FING22_value,
                                    'FING23_value': FING23_value,
                                    'FING24_value': FING24_value,
                                    }
                FING_Max_value = max(All_Fingers_Dict.values())
                FING_String = ''
                for key, value in All_Fingers_Dict.items():
                    if value == FING_Max_value:
                        FING_String = key
                        self.lines.append(int(FING_String[4:6]))
                        value_3_digits = round(value, 3) # 保留三位有效数字
                        self.lines.append(value_3_digits)
                        break

                # 求个平均值
                normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                index1 = np.where(abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                index1 = index1[0][0]
                FING_value1 = self.log.data[FING_String[0:6]][index1]

                index2 = np.where(abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                index2 = index2[0][0]
                FING_value2 = self.log.data[FING_String[0:6]][index2]

                index3 = np.where(abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                index3 = index3[0][0]
                FING_value3 = self.log.data[FING_String[0:6]][index3]

                index4 = np.where(abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                index4 = index4[0][0]
                FING_value4 = self.log.data[FING_String[0:6]][index4]

                index5 = np.where(abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                index5 = index5[0][0]
                FING_value5 = self.log.data[FING_String[0:6]][index5]

                FING_value = round(sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                self.lines.append(FING_value)
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Penetration.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()
            elif self.damage_Tag == 'Projection':
                All_Fingers = [FING01_value, FING02_value, FING03_value, FING04_value, FING05_value, FING06_value,
                               FING07_value,
                               FING08_value, FING09_value, FING10_value, FING11_value, FING12_value, FING13_value,
                               FING14_value,
                               FING15_value, FING16_value, FING17_value, FING18_value, FING19_value, FING20_value,
                               FING21_value,
                               FING22_value, FING23_value, FING24_value]
                All_Fingers_Dict = {'FING01_value': FING01_value,
                                    'FING02_value': FING02_value,
                                    'FING03_value': FING03_value,
                                    'FING04_value': FING04_value,
                                    'FING05_value': FING05_value,
                                    'FING06_value': FING06_value,
                                    'FING07_value': FING07_value,
                                    'FING08_value': FING08_value,
                                    'FING09_value': FING09_value,
                                    'FING10_value': FING10_value,
                                    'FING11_value': FING11_value,
                                    'FING12_value': FING12_value,
                                    'FING13_value': FING13_value,
                                    'FING14_value': FING14_value,
                                    'FING15_value': FING15_value,
                                    'FING16_value': FING16_value,
                                    'FING17_value': FING17_value,
                                    'FING18_value': FING18_value,
                                    'FING19_value': FING19_value,
                                    'FING20_value': FING20_value,
                                    'FING21_value': FING21_value,
                                    'FING22_value': FING22_value,
                                    'FING23_value': FING23_value,
                                    'FING24_value': FING24_value,
                                    }
                FING_Min_value = min(All_Fingers_Dict.values())
                FING_String = ''
                for key, value in All_Fingers_Dict.items():
                    if value == FING_Min_value:
                        FING_String = key
                        self.lines.append(int(FING_String[4:6]))
                        value_3_digits = round(value, 3)  # 保留三位有效数字
                        self.lines.append(value_3_digits)
                        break

                # 求个平均值
                normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                index1 = np.where(abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                index1 = index1[0][0]
                FING_value1 = self.log.data[FING_String[0:6]][index1]

                index2 = np.where(abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                index2 = index2[0][0]
                FING_value2 = self.log.data[FING_String[0:6]][index2]

                index3 = np.where(abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                index3 = index3[0][0]
                FING_value3 = self.log.data[FING_String[0:6]][index3]

                index4 = np.where(abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                index4 = index4[0][0]
                FING_value4 = self.log.data[FING_String[0:6]][index4]

                index5 = np.where(abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                index5 = index5[0][0]
                FING_value5 = self.log.data[FING_String[0:6]][index5]

                FING_value = round(sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                self.lines.append(FING_value)
                self.lines.append(round(self.Min_value, 3))
                self.lines.append(round(self.Ave_value, 3))
                self.lines.append(round(self.Max_value, 3))
                print(self.lines)
                # 保存list到文件
                self.lines_temp = pd.DataFrame(self.lines, columns=['value'])
                writer = pd.ExcelWriter('Projection.xlsx')
                self.lines_temp.to_excel(writer, 'Sheet1')
                writer.save()
            elif self.damage_Tag == 'Holes-detection':
                number = int((float(self.lines[1]) - float(self.lines[0])) / 0.01)
                for i in range(number):
                    current_depth = abs(float(self.lines[0]) + 0.01*i)
                    # print("当前深度为: " + str(current_depth) + "m")
                    index = np.where(abs(self.log.data['DEPT'] - (float(self.lines[0]) + 0.01*i)) <= float(abs(self.log.step) - 0.001))
                    index = index[0][0]
                    FING01_value = self.log.data['FING01'][index]
                    FING02_value = self.log.data['FING02'][index]
                    FING03_value = self.log.data['FING03'][index]
                    FING04_value = self.log.data['FING04'][index]
                    FING05_value = self.log.data['FING05'][index]
                    FING06_value = self.log.data['FING06'][index]
                    FING07_value = self.log.data['FING07'][index]
                    FING08_value = self.log.data['FING08'][index]
                    FING09_value = self.log.data['FING09'][index]
                    FING10_value = self.log.data['FING10'][index]
                    FING11_value = self.log.data['FING11'][index]
                    FING12_value = self.log.data['FING12'][index]
                    FING13_value = self.log.data['FING13'][index]
                    FING14_value = self.log.data['FING14'][index]
                    FING15_value = self.log.data['FING15'][index]
                    FING16_value = self.log.data['FING16'][index]
                    FING17_value = self.log.data['FING17'][index]
                    FING18_value = self.log.data['FING18'][index]
                    FING19_value = self.log.data['FING19'][index]
                    FING20_value = self.log.data['FING20'][index]
                    FING21_value = self.log.data['FING21'][index]
                    FING22_value = self.log.data['FING22'][index]
                    FING23_value = self.log.data['FING23'][index]
                    FING24_value = self.log.data['FING24'][index]
                    self.Min_value = self.log.data['MINDIA'][index]
                    self.Ave_value = self.log.data['AVEDIA'][index]
                    self.Max_value = self.log.data['MAXDIA'][index]
                    # print(str(self.Min_value), ' ', str(self.Ave_value), ' ', str(self.Max_value))

                    All_Fingers = [FING01_value, FING02_value, FING03_value, FING04_value, FING05_value, FING06_value,
                                   FING07_value,
                                   FING08_value, FING09_value, FING10_value, FING11_value, FING12_value, FING13_value,
                                   FING14_value,
                                   FING15_value, FING16_value, FING17_value, FING18_value, FING19_value, FING20_value,
                                   FING21_value,
                                   FING22_value, FING23_value, FING24_value]
                    All_Fingers_Dict = {'FING01_value': FING01_value,
                                        'FING02_value': FING02_value,
                                        'FING03_value': FING03_value,
                                        'FING04_value': FING04_value,
                                        'FING05_value': FING05_value,
                                        'FING06_value': FING06_value,
                                        'FING07_value': FING07_value,
                                        'FING08_value': FING08_value,
                                        'FING09_value': FING09_value,
                                        'FING10_value': FING10_value,
                                        'FING11_value': FING11_value,
                                        'FING12_value': FING12_value,
                                        'FING13_value': FING13_value,
                                        'FING14_value': FING14_value,
                                        'FING15_value': FING15_value,
                                        'FING16_value': FING16_value,
                                        'FING17_value': FING17_value,
                                        'FING18_value': FING18_value,
                                        'FING19_value': FING19_value,
                                        'FING20_value': FING20_value,
                                        'FING21_value': FING21_value,
                                        'FING22_value': FING22_value,
                                        'FING23_value': FING23_value,
                                        'FING24_value': FING24_value,
                                        }
                    FING_Max_value = max(All_Fingers_Dict.values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict.items():
                        if value == FING_Max_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6])) # 获取测量臂编号
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break

                    # 求个平均值
                    normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                    normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                    normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                    normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                    normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                    index1 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                    index1 = index1[0][0]
                    FING_value1 = self.log.data[FING_String[0:6]][index1]

                    index2 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                    index2 = index2[0][0]
                    FING_value2 = self.log.data[FING_String[0:6]][index2]

                    index3 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                    index3 = index3[0][0]
                    FING_value3 = self.log.data[FING_String[0:6]][index3]

                    index4 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                    index4 = index4[0][0]
                    FING_value4 = self.log.data[FING_String[0:6]][index4]

                    index5 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                    index5 = index5[0][0]
                    FING_value5 = self.log.data[FING_String[0:6]][index5]

                    FING_value = round(
                        sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                    self.lines.append(FING_value)
                    self.lines.append(round(self.Min_value, 3))
                    self.lines.append(round(self.Ave_value, 3))
                    self.lines.append(round(self.Max_value, 3))

                    ################################################### 评价

                    data = pd.DataFrame({'损伤井段(m)': '',
                                         '最大损伤点深度(m)': '',
                                         '单臂测量值(mm)': '',
                                         '最小内径(mm)': '',
                                         '平均内径(mm)': '',
                                         '最大内径(mm)': '',
                                         '损伤量(mm)': '',
                                         '损伤程度(%)': '',
                                         '损伤级别': ''},
                                        index=[1])

                    well_interval = ''.join([str(self.lines[0]), '-', str(self.lines[1])])
                    penetration_value = float(self.lines[4] - float(self.lines[5]))
                    penetration_value = round(penetration_value, 2)
                    penetration_degree = penetration_value / float(self.thickness.value) * 100
                    penetration_degree = round(penetration_degree, 2)

                    # 将某一损伤阈值的深度全部输出
                    #TODO
                    if penetration_degree > 30: # 损伤阈值参数可调
                        # 损伤级别判定
                        penetration_describe = ''
                        if penetration_degree < 10:
                            penetration_describe = '一级损伤'
                            print('损伤程度: ', penetration_degree, '%: 一级损伤')
                        elif 10 <= penetration_degree < 20:
                            penetration_describe = '二级损伤'
                            print('损伤程度: ', penetration_degree, '%: 二级损伤')
                        elif 20 <= penetration_degree < 40:
                            penetration_describe = '三级损伤'
                            print('损伤程度: ', penetration_degree, '%: 三级损伤')
                        elif 40 <= penetration_degree < 85:
                            penetration_describe = '四级损伤'
                            print('损伤程度: ', penetration_degree, '%: 四级损伤')
                        elif penetration_degree >= 85:
                            penetration_describe = '五级损伤'
                            print('损伤程度: ', penetration_degree, '%: 五级损伤')
                        print("当前深度为: " + str(round(current_depth, 2)) + "m")
                        print(str(self.lines) + '\n')

                    self.lines = self.lines[0:3] # 保留前三个值，后面的删除 eg:[3435.23, 3453.92, 3445.12, 2, 58.672, 57.519, 112.427, 114.442, 116.678]

                    ######################################
                    # 剩下第二大的值
                    All_Fingers_Dict.pop(FING_String)
                    # print(All_Fingers_Dict)
                    FING_Max_value = max(All_Fingers_Dict.values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict.items():
                        if value == FING_Max_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6]))  # 获取测量臂编号
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break
                    # 求个平均值
                    normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                    normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                    normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                    normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                    normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                    index1 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                    index1 = index1[0][0]
                    FING_value1 = self.log.data[FING_String[0:6]][index1]

                    index2 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                    index2 = index2[0][0]
                    FING_value2 = self.log.data[FING_String[0:6]][index2]

                    index3 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                    index3 = index3[0][0]
                    FING_value3 = self.log.data[FING_String[0:6]][index3]

                    index4 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                    index4 = index4[0][0]
                    FING_value4 = self.log.data[FING_String[0:6]][index4]

                    index5 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                    index5 = index5[0][0]
                    FING_value5 = self.log.data[FING_String[0:6]][index5]

                    FING_value = round(
                        sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                    self.lines.append(FING_value)
                    self.lines.append(round(self.Min_value, 3))
                    self.lines.append(round(self.Ave_value, 3))
                    self.lines.append(round(self.Max_value, 3))

                    ################################################### 评价

                    data = pd.DataFrame({'损伤井段(m)': '',
                                         '最大损伤点深度(m)': '',
                                         '单臂测量值(mm)': '',
                                         '最小内径(mm)': '',
                                         '平均内径(mm)': '',
                                         '最大内径(mm)': '',
                                         '损伤量(mm)': '',
                                         '损伤程度(%)': '',
                                         '损伤级别': ''},
                                        index=[1])

                    well_interval = ''.join([str(self.lines[0]), '-', str(self.lines[1])])
                    penetration_value = float(self.lines[4] - float(self.lines[5]))
                    penetration_value = round(penetration_value, 2)
                    penetration_degree = penetration_value / float(self.thickness.value) * 100
                    penetration_degree = round(penetration_degree, 2)

                    # 将某一损伤阈值的深度全部输出
                    # TODO
                    if penetration_degree > 30:  # 损伤阈值参数可调
                        # 损伤级别判定
                        penetration_describe = ''
                        if penetration_degree < 10:
                            penetration_describe = '一级损伤'
                            print('损伤程度: ', penetration_degree, '%: 一级损伤')
                        elif 10 <= penetration_degree < 20:
                            penetration_describe = '二级损伤'
                            print('损伤程度: ', penetration_degree, '%: 二级损伤')
                        elif 20 <= penetration_degree < 40:
                            penetration_describe = '三级损伤'
                            print('损伤程度: ', penetration_degree, '%: 三级损伤')
                        elif 40 <= penetration_degree < 85:
                            penetration_describe = '四级损伤'
                            print('损伤程度: ', penetration_degree, '%: 四级损伤')
                        elif penetration_degree >= 85:
                            penetration_describe = '五级损伤'
                            print('损伤程度: ', penetration_degree, '%: 五级损伤')
                        print("当前深度为: " + str(round(current_depth, 2)) + "m")
                        print(str(self.lines) + '\n')

                    self.lines = self.lines[0:3]
                    ######################################
                    # 剩下第三大的值
                    All_Fingers_Dict.pop(FING_String)
                    # print(All_Fingers_Dict)
                    FING_Max_value = max(All_Fingers_Dict.values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict.items():
                        if value == FING_Max_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6]))  # 获取测量臂编号
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break
                    # 求个平均值
                    normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                    normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                    normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                    normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                    normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                    index1 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                    index1 = index1[0][0]
                    FING_value1 = self.log.data[FING_String[0:6]][index1]

                    index2 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                    index2 = index2[0][0]
                    FING_value2 = self.log.data[FING_String[0:6]][index2]

                    index3 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                    index3 = index3[0][0]
                    FING_value3 = self.log.data[FING_String[0:6]][index3]

                    index4 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                    index4 = index4[0][0]
                    FING_value4 = self.log.data[FING_String[0:6]][index4]

                    index5 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                    index5 = index5[0][0]
                    FING_value5 = self.log.data[FING_String[0:6]][index5]

                    FING_value = round(
                        sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                    self.lines.append(FING_value)
                    self.lines.append(round(self.Min_value, 3))
                    self.lines.append(round(self.Ave_value, 3))
                    self.lines.append(round(self.Max_value, 3))

                    ################################################### 评价

                    data = pd.DataFrame({'损伤井段(m)': '',
                                         '最大损伤点深度(m)': '',
                                         '单臂测量值(mm)': '',
                                         '最小内径(mm)': '',
                                         '平均内径(mm)': '',
                                         '最大内径(mm)': '',
                                         '损伤量(mm)': '',
                                         '损伤程度(%)': '',
                                         '损伤级别': ''},
                                        index=[1])

                    well_interval = ''.join([str(self.lines[0]), '-', str(self.lines[1])])
                    penetration_value = float(self.lines[4] - float(self.lines[5]))
                    penetration_value = round(penetration_value, 2)
                    penetration_degree = penetration_value / float(self.thickness.value) * 100
                    penetration_degree = round(penetration_degree, 2)

                    # 将某一损伤阈值的深度全部输出
                    # TODO
                    if penetration_degree > 30:  # 损伤阈值参数可调
                        # 损伤级别判定
                        penetration_describe = ''
                        if penetration_degree < 10:
                            penetration_describe = '一级损伤'
                            print('损伤程度: ', penetration_degree, '%: 一级损伤')
                        elif 10 <= penetration_degree < 20:
                            penetration_describe = '二级损伤'
                            print('损伤程度: ', penetration_degree, '%: 二级损伤')
                        elif 20 <= penetration_degree < 40:
                            penetration_describe = '三级损伤'
                            print('损伤程度: ', penetration_degree, '%: 三级损伤')
                        elif 40 <= penetration_degree < 85:
                            penetration_describe = '四级损伤'
                            print('损伤程度: ', penetration_degree, '%: 四级损伤')
                        elif penetration_degree >= 85:
                            penetration_describe = '五级损伤'
                            print('损伤程度: ', penetration_degree, '%: 五级损伤')
                        print("当前深度为: " + str(round(current_depth, 2)) + "m")
                        print(str(self.lines) + '\n')

                    self.lines = self.lines[0:3]
                    ######################################
                    # 剩下第四大的值
                    All_Fingers_Dict.pop(FING_String)
                    # print(All_Fingers_Dict)
                    FING_Max_value = max(All_Fingers_Dict.values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict.items():
                        if value == FING_Max_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6]))  # 获取测量臂编号
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break
                    # 求个平均值
                    normal_Depth1 = float(self.lines[0]) - 0.2  # 在井段开始深度上方0.2m处取一个正常的点进行读值
                    normal_Depth2 = float(self.lines[0]) - 0.4  # 在井段开始深度上方0.4m处取一个正常的点进行读值
                    normal_Depth3 = float(self.lines[0]) - 0.6  # 在井段开始深度上方0.6m处取一个正常的点进行读值
                    normal_Depth4 = float(self.lines[0]) - 0.8  # 在井段开始深度上方0.8m处取一个正常的点进行读值
                    normal_Depth5 = float(self.lines[0]) - 1.0  # 在井段开始深度上方1.0m处取一个正常的点进行读值

                    index1 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth1) <= float(abs(self.log.step) - 0.001))
                    index1 = index1[0][0]
                    FING_value1 = self.log.data[FING_String[0:6]][index1]

                    index2 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth2) <= float(abs(self.log.step) - 0.001))
                    index2 = index2[0][0]
                    FING_value2 = self.log.data[FING_String[0:6]][index2]

                    index3 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth3) <= float(abs(self.log.step) - 0.001))
                    index3 = index3[0][0]
                    FING_value3 = self.log.data[FING_String[0:6]][index3]

                    index4 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth4) <= float(abs(self.log.step) - 0.001))
                    index4 = index4[0][0]
                    FING_value4 = self.log.data[FING_String[0:6]][index4]

                    index5 = np.where(
                        abs(self.log.data['DEPT'] - normal_Depth5) <= float(abs(self.log.step) - 0.001))
                    index5 = index5[0][0]
                    FING_value5 = self.log.data[FING_String[0:6]][index5]

                    FING_value = round(
                        sum([FING_value1, FING_value2, FING_value3, FING_value4, FING_value5]) / 5.0, 3)

                    self.lines.append(FING_value)
                    self.lines.append(round(self.Min_value, 3))
                    self.lines.append(round(self.Ave_value, 3))
                    self.lines.append(round(self.Max_value, 3))

                    ################################################### 评价

                    data = pd.DataFrame({'损伤井段(m)': '',
                                         '最大损伤点深度(m)': '',
                                         '单臂测量值(mm)': '',
                                         '最小内径(mm)': '',
                                         '平均内径(mm)': '',
                                         '最大内径(mm)': '',
                                         '损伤量(mm)': '',
                                         '损伤程度(%)': '',
                                         '损伤级别': ''},
                                        index=[1])

                    well_interval = ''.join([str(self.lines[0]), '-', str(self.lines[1])])
                    penetration_value = float(self.lines[4] - float(self.lines[5]))
                    penetration_value = round(penetration_value, 2)
                    penetration_degree = penetration_value / float(self.thickness.value) * 100
                    penetration_degree = round(penetration_degree, 2)

                    # 将某一损伤阈值的深度全部输出
                    # TODO
                    if penetration_degree > 30:  # 损伤阈值参数可调
                        # 损伤级别判定
                        penetration_describe = ''
                        if penetration_degree < 10:
                            penetration_describe = '一级损伤'
                            print('损伤程度: ', penetration_degree, '%: 一级损伤')
                        elif 10 <= penetration_degree < 20:
                            penetration_describe = '二级损伤'
                            print('损伤程度: ', penetration_degree, '%: 二级损伤')
                        elif 20 <= penetration_degree < 40:
                            penetration_describe = '三级损伤'
                            print('损伤程度: ', penetration_degree, '%: 三级损伤')
                        elif 40 <= penetration_degree < 85:
                            penetration_describe = '四级损伤'
                            print('损伤程度: ', penetration_degree, '%: 四级损伤')
                        elif penetration_degree >= 85:
                            penetration_describe = '五级损伤'
                            print('损伤程度: ', penetration_degree, '%: 五级损伤')
                        print("当前深度为: " + str(round(current_depth, 2)) + "m")
                        print(str(self.lines) + '\n')

                    self.lines = self.lines[0:3]
                print('===============================================================================')
            elif self.damage_Tag == 'Type-suggestion':
                number = int((float(self.lines[1]) - float(self.lines[0])) / 0.01)
                All_Fingers = {}
                All_Fingers_Dict = {}
                All_Diameter_Dict = {}

                for i in range(number):
                    current_depth = abs(float(self.lines[0]) + 0.01*i)
                    # print("当前深度为: " + str(current_depth) + "m")
                    index = np.where(abs(self.log.data['DEPT'] - (float(self.lines[0]) + 0.01*i)) <= float(abs(self.log.step) - 0.001))
                    index = index[0][0]
                    FING01_value = self.log.data['FING01'][index]
                    FING02_value = self.log.data['FING02'][index]
                    FING03_value = self.log.data['FING03'][index]
                    FING04_value = self.log.data['FING04'][index]
                    FING05_value = self.log.data['FING05'][index]
                    FING06_value = self.log.data['FING06'][index]
                    FING07_value = self.log.data['FING07'][index]
                    FING08_value = self.log.data['FING08'][index]
                    FING09_value = self.log.data['FING09'][index]
                    FING10_value = self.log.data['FING10'][index]
                    FING11_value = self.log.data['FING11'][index]
                    FING12_value = self.log.data['FING12'][index]
                    FING13_value = self.log.data['FING13'][index]
                    FING14_value = self.log.data['FING14'][index]
                    FING15_value = self.log.data['FING15'][index]
                    FING16_value = self.log.data['FING16'][index]
                    FING17_value = self.log.data['FING17'][index]
                    FING18_value = self.log.data['FING18'][index]
                    FING19_value = self.log.data['FING19'][index]
                    FING20_value = self.log.data['FING20'][index]
                    FING21_value = self.log.data['FING21'][index]
                    FING22_value = self.log.data['FING22'][index]
                    FING23_value = self.log.data['FING23'][index]
                    FING24_value = self.log.data['FING24'][index]
                    self.Min_value = self.log.data['MINDIA'][index]
                    self.Ave_value = self.log.data['AVEDIA'][index]
                    self.Max_value = self.log.data['MAXDIA'][index]
                    # print(str(self.Min_value), ' ', str(self.Ave_value), ' ', str(self.Max_value))
                    All_Fingers[i] = [FING01_value, FING02_value, FING03_value, FING04_value, FING05_value, FING06_value,
                                   FING07_value,
                                   FING08_value, FING09_value, FING10_value, FING11_value, FING12_value, FING13_value,
                                   FING14_value,
                                   FING15_value, FING16_value, FING17_value, FING18_value, FING19_value, FING20_value,
                                   FING21_value,
                                   FING22_value, FING23_value, FING24_value]
                    All_Fingers_Dict[i] = {'FING01_value': FING01_value,
                                        'FING02_value': FING02_value,
                                        'FING03_value': FING03_value,
                                        'FING04_value': FING04_value,
                                        'FING05_value': FING05_value,
                                        'FING06_value': FING06_value,
                                        'FING07_value': FING07_value,
                                        'FING08_value': FING08_value,
                                        'FING09_value': FING09_value,
                                        'FING10_value': FING10_value,
                                        'FING11_value': FING11_value,
                                        'FING12_value': FING12_value,
                                        'FING13_value': FING13_value,
                                        'FING14_value': FING14_value,
                                        'FING15_value': FING15_value,
                                        'FING16_value': FING16_value,
                                        'FING17_value': FING17_value,
                                        'FING18_value': FING18_value,
                                        'FING19_value': FING19_value,
                                        'FING20_value': FING20_value,
                                        'FING21_value': FING21_value,
                                        'FING22_value': FING22_value,
                                        'FING23_value': FING23_value,
                                        'FING24_value': FING24_value,
                                        }
                    FING_Max_value = max(All_Fingers_Dict[i].values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict[i].items():
                        if value == FING_Max_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6]))
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break

                    FING_Min_value = min(All_Fingers_Dict[i].values())
                    FING_String = ''
                    for key, value in All_Fingers_Dict[i].items():
                        if value == FING_Min_value:
                            FING_String = key
                            self.lines.append(int(FING_String[4:6]))
                            value_3_digits = round(value, 3)  # 保留三位有效数字
                            self.lines.append(value_3_digits)
                            break
                    # 当前lines为[3435.23, 3453.92, 3445.12, 2, 58.672, 25, 56.468]
                    # print('i=', i, ' ', All_Fingers_Dict[i])

                    All_Diameter_Dict[i] = {'Diameter01_value': round(FING01_value + FING13_value, 3),
                                           'Diameter02_value': round(FING02_value + FING14_value, 3),
                                           'Diameter03_value': round(FING03_value + FING15_value, 3),
                                           'Diameter04_value': round(FING04_value + FING16_value, 3),
                                           'Diameter05_value': round(FING05_value + FING17_value, 3),
                                           'Diameter06_value': round(FING06_value + FING18_value, 3),
                                           'Diameter07_value': round(FING07_value + FING19_value, 3),
                                           'Diameter08_value': round(FING08_value + FING20_value, 3),
                                           'Diameter09_value': round(FING09_value + FING21_value, 3),
                                           'Diameter10_value': round(FING10_value + FING22_value, 3),
                                           'Diameter11_value': round(FING11_value + FING23_value, 3),
                                           'Diameter12_value': round(FING12_value + FING24_value, 3),
                                           }


                    self.lines = self.lines[0:3] # 保留前三个值，后面的删除 eg:[3435.23, 3453.92, 3445.12, 2, 58.672, 57.519, 112.427, 114.442, 116.678]
                print(All_Fingers_Dict)
                print('------------')
                print(All_Diameter_Dict)

                # 获取正常段的内径测量值
                index = np.where(abs(self.log.data['DEPT'] - float(self.lines[1])) <= float(abs(self.log.step) - 0.001))
                index = index[0][0]
                Normal_value = self.log.data['AVEDIA'][index]
                print('Normal value is: ', Normal_value)

                count1 = 0
                count2 = 0
                all_num = i
                for num in range(all_num + 1):
                    for key, value in All_Diameter_Dict[num].items():
                        if value - float(Normal_value) > 2: #TODO
                            count1 = count1 + 1
                        if value - float(Normal_value) < -2:
                            count2 = count2 + 1
                # 最小内径平均值大于标准内径的2mm以上的比例
                ratio1 = count1 / (all_num * 24) * 100
                # 最小内径平均值小于标准内径的2mm以上的比例
                ratio2 = count2 / (all_num * 24) * 100
                print('最小内径平均值大于标准内径的2mm以上的比例: ', ratio1)
                print('最小内径平均值小于标准内径的2mm以上的比例: ', ratio2)
                if ratio1 > 10 and ratio2 > 10:
                    print('该处可能存在挤压变形')
                elif ratio1 > 10 and ratio2 <= 10:
                    print('该处可能存在扩径/错位')
                elif ratio1 <= 10 and ratio2 > 10:
                    print('该处可能存在缩径/错位')
                elif 0 <= ratio1 < 2 and ratio2 <= 0.1:
                    print('该处可能存在点状损伤')
                elif 0 <= ratio2 < 2 and ratio1 <= 0.1:
                    print('该处可能存在点状结垢')
                else:
                    print('暂时无法识别损伤类型')

            self.fig2 = plt.figure('截面图')
            # 设置下面所需要的参数
            barSlices1 = 24
            barSlices2 = 100

            # theta指每个标记所在射线与极径的夹角，下面表示均分角度
            theta1 = np.linspace(0.0, 2 * np.pi, barSlices1, endpoint=False)
            theta2 = np.linspace(0.0, 2 * np.pi, barSlices2, endpoint=False)

            # # r表示点距离圆心的距离，np.random.rand(barSlices)表示返回返回服从“0-1”均匀分布的随机样本值
            # r = 2 * np.random.rand(barSlices) + 50
            r = [FING01_value, FING02_value, FING03_value, FING04_value, FING05_value, FING06_value, FING07_value,
                 FING08_value, FING09_value, FING10_value, FING11_value, FING12_value, FING13_value, FING14_value,
                 FING15_value, FING16_value, FING17_value, FING18_value, FING19_value, FING20_value, FING21_value,
                 FING22_value, FING23_value, FING24_value]

            # 网上搜的方法，不知道怎么就可以闭合了(黑人问号)
            r = np.concatenate((r, [r[0]]))  # 闭合
            theta1 = np.concatenate((theta1, [theta1[0]]))  # 闭合
            theta2 = np.concatenate((theta2, [theta2[0]]))  # 闭合

            inside_radius = [float(self.inner_diameter.value) / 2] * barSlices2
            outside_radius = [float(self.outer_diameter.value) / 2] * barSlices2

            inside_radius = np.concatenate((inside_radius, [inside_radius[0]]))  # 闭合
            outside_radius = np.concatenate((outside_radius, [outside_radius[0]]))  # 闭合

            # 绘图之前先清理一下
            plt.clf()
            # polar表示绘制极坐标图，颜色，线宽，标志点样式
            plt.polar(theta1, r, color="blue", linewidth=1, marker="", mfc="b", ms=10)
            plt.gca().set_theta_zero_location('N')
            # plt.gca().set_rlim(0, 57.15)  # 设置显示的极径范围
            plt.gca().set_rlim(0, float(self.outer_diameter.value) / 2)
            plt.gca().fill(theta1, r, facecolor='w', alpha=0.2)  # 填充颜色
            plt.gca().fill_between(theta2, outside_radius, inside_radius, facecolor='cyan', alpha=0.8)  # 填充之间颜色
            plt.gca().patch.set_facecolor('1')
            plt.gca().set_rgrids(np.arange(0, 100, 100))

            # label = np.array([j for j in range(1, 25)])  # 定义标签
            # plt.gca().set_thetagrids(np.arange(0, 360, 15), label)
            plt.gca().set_thetagrids(np.arange(0, 360, 360))

            text_show = ''.join(['Depth:', str(ymax), 'm\nOD: ', str(self.outer_diameter.value), 'mm\nID: ', \
                                 str(self.inner_diameter.value),'mm\nWT: ', str(self.thickness.value), 'mm'])
            plt.text(7 * np.pi / 4, 90, text_show)
            # 绘图展示
            plt.show()

    # RadioButtons行为定义
    def actionfunc(self, damage_type):

        # 损伤SpanSelector对象
        self.span3_red = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                      rectprops=dict(alpha=0.5, facecolor='red'), span_stays=True)
        # 结垢SpanSelector对象
        self.span3_green = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                        rectprops=dict(alpha=0.5, facecolor='green'), span_stays=True)
        # 变形SpanSelector对象
        self.span3_yellow = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                         rectprops=dict(alpha=0.5, facecolor='yellow'), span_stays=True)

        self.span3_orange = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                         rectprops=dict(alpha=0.5, facecolor='orange'), span_stays=True)

        self.span3_blue = SpanSelector(self.ax3, self.onselect3, 'vertical', useblit=True,
                                         rectprops=dict(alpha=0.5, facecolor='blue'), span_stays=True)

        if damage_type == 'Penetration':
            print(damage_type)
            self.damage_Tag = 'Penetration'
            # del self.span3_red
            del self.span3_green
            del self.span3_yellow
            del self.span3_orange
            del self.span3_blue
        elif damage_type == 'Projection':
            print(damage_type)
            self.damage_Tag = 'Projection'
            del self.span3_red
            # del self.span3_green
            del self.span3_yellow
            del self.span3_orange
            del self.span3_blue
        elif damage_type == 'Transformation':
            print(damage_type)
            self.damage_Tag = 'Transformation'
            del self.span3_red
            del self.span3_green
            # del self.span3_yellow
            del self.span3_orange
            del self.span3_blue
        elif damage_type == 'Holes-detection':
            print(damage_type)
            self.damage_Tag = 'Holes-detection'
            del self.span3_red
            del self.span3_green
            del self.span3_yellow
            # del self.span3_orange
            del self.span3_blue
        elif damage_type == 'Type-suggestion':
            print(damage_type)
            self.damage_Tag = 'Type-suggestion'
            del self.span3_red
            del self.span3_green
            del self.span3_yellow
            del self.span3_orange
            # del self.span3_blue

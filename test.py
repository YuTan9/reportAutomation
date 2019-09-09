# pack to exe: powershell run: auto-py-to-exe
# one file is not recommended, use one directory instead.
# include the additional files on the run (icon.png)
from appJar import gui
import fetchGoogleSheet as fgs
import pdfkit
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import threading
import time

def getLineChart(x,y):
	# x = ['2018-10-11','2018-10-12', '2018-10-13','2018-10-14','2018-10-15','2018-10-16']
	# y = [1,3,5,7,13,7]
	# if(len(x) < 6):
	x = ['2019/5/06', '2019/7/25']
	pseudoX = x
	y = [6, 1]
	# while len(pseudoX) < 6:
	# 	pseudoX.append(' ')
	# 	y.append(y[len(y)-1])
	if len(pseudoX) < 6:
		pseudoX.append(' ')
		y.append(y[len(y)-1])

	plt.style.use('fast')
	fig = plt.figure(figsize = [7,7])
	ax = plt.axes()
	minimum = min(y)
	maximum = max(y)
	range_  = maximum - minimum
	ticks   = int(maximum / 8)+1 # a $ticks between every ticks

	ytick   = []
	for i in range(8):
		ytick.append(i*ticks)

	plt.xticks(range(len(pseudoX)), pseudoX)
	plt.yticks(ytick, ytick)
	plt.xlabel("Test Date")
	plt.ylabel("CTC Count")

	y0= [ytick[0] for i in range(len(y))]
	y1= [ytick[1] for i in range(len(y))]
	y2= [ytick[2] for i in range(len(y))]
	y3= [ytick[3] for i in range(len(y))]
	y4= [ytick[4] for i in range(len(y))]
	y5= [ytick[5] for i in range(len(y))]
	y6= [ytick[6] for i in range(len(y))]
	y7= [ytick[7] for i in range(len(y))]
	plt.plot(pseudoX,y0, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y1, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y2, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y3, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y4, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y5, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y6, '-', color='black', linewidth=1)
	plt.plot(pseudoX,y7, '-', color='black', linewidth=1)

	base = [y[0] for i in range(len(y))]
	plt.plot(pseudoX, base, '-', color='red', linewidth=2)

	plt.plot(['2019/5/06', '2019/7/25'],[6,1], '-', color='black', linewidth=2, markersize=1)
	plt.plot(x[0], y[0], 'o', color='black')

	
	fig.savefig("tmp.jpg",figsize=(7, 7), dpi=100)
	# plt.show()
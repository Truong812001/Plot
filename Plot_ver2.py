import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import psse35
class Plot():
	def __init__(self,fi):
		self.BUS=pd.read_excel(fi,sheet_name='BUS',skiprows=1)
		self.LINE=pd.read_excel(fi,sheet_name='LINE',skiprows=1)
		self.coordinates()

	def coordinates(self):

		## tọa độ bus 
		self.coor_bus = dict()

		busID = self.BUS['NO']
		x=self.BUS['x']
		y=self.BUS['y']

		for i in range(len(busID)):
			a=[]
			a.append(x[i])
			a.append(y[i])
			self.coor_bus[busID[i]] = a


		self.line = dict()
		lineID=self.LINE['NO']
		frombus=self.LINE['FROMBUS']
		tobus=self.LINE['TOBUS']

		for i in range(len(lineID)):
			a=[]
			a.append(frombus[i])
			a.append(tobus[i])
			self.line[lineID[i]] = a


		return
	def main(self):

		print(self.coor_bus)

		special_x=self.LINE['x']
		special_y=self.LINE['y']

		special_line={}

		for i in range(len(self.LINE['NO'])):
			if pd.isna(special_x[i]) == False:
				xadd=str(special_x[i]).split()
				yadd=str(special_y[i]).split()
				for j in range(len(xadd)):

					special_line.setdefault(self.LINE['NO'][i], []).append([float(xadd[j]),float(yadd[j])])




		plt.figure(figsize=(10, 6))

		## draw point
		for bus in self.coor_bus.keys():
			x=self.coor_bus[bus][0]
			y=self.coor_bus[bus][1]
			plt.scatter(x, y,color='black')
			plt.annotate(bus,(x+0.1,y+0.02),fontsize=size)


		## Flag
		flag=self.LINE['FLAG']
		name_line=self.LINE['NO']
		nline_off=[37, 7, 9, 14, 28]

		for i in range(len(flag)):
			for li in nline_off :
				if name_line[i]==li:
					flag[i]=0



		## draw line 
		for i,li in enumerate(self.line.keys()):

			## đường dây 
			if li not in special_line:
				x=[]
				y=[]
				for bus in self.line[li]:
					x.append(self.coor_bus[bus][0])
					y.append(self.coor_bus[bus][1])
				print(x,y)
				## Name Line 
				self.draw_name_line(x,y,li,i)

				self.plot(x,y,flag,i,li)
				# self.rate_line(bus,li)
			#đường dây gấp khúc 
			else:

				for j,spec in enumerate(special_line[li]):

					## bus đầu 
					if j == 0 :
			
						bus=self.line[li][0]

						x=[self.coor_bus[bus][0],spec[0]]
						y=[self.coor_bus[bus][1],spec[1]]
						self.plot(x,y,flag,i,li)
						## name_line
						self.draw_name_line(x,y,li,i)

					##bus cuối 
					if j==len(special_line[li])-1:
					
						bus=self.line[li][1]
						x=[spec[0],self.coor_bus[bus][0]]
						y=[spec[1],self.coor_bus[bus][1]]
						self.plot(x,y,flag,i,li)
					##bus trung gian
					else:
						x=[special_line[li][j][0],special_line[li][j+1][0]]
						y=[special_line[li][j][1],special_line[li][j+1][1]]
						self.plot(x,y,flag,i,li)

		# x=[1,3]
		# y=[0,0]

		# plt.plot(x, y,'k', linestyle='solid',color='r')




	def plot(self,x,y,flag,i,li):
		line_off=()

		color='black'
		if li in line_off:
			plt.plot(x, y, linestyle='--',color=color)
		elif flag[i]==1:
			plt.plot(x, y, linestyle='solid',color=color)
		elif flag[i]==0:
			plt.plot(x, y, linestyle='--',color=color)
	def draw_name_line(self,x,y,li,i):

		# tọa độ x bằng nhau
		if x[0]==x[1] :
			x1=x[0]+0.25
			y1=(y[0]+y[1])/2
			plt.annotate(li,(x1,y1),fontsize=size,fontstyle='italic')

			plt.plot([x1-0.02,x1+0.3],[y1-0.03,y1-0.03],'k',linestyle='solid')
			print('ok',li)
		## tọa dộ y bằng nhau
		if y[0]==y[1]:
			x1=(x[0]+x[1])/2
			y1=y[1]+0.06
			plt.annotate(li,(x1,y1),fontsize=size,fontstyle='italic')

			plt.plot([x1-0.02,x1+0.3],[y1-0.03,y1-0.03],'k',linestyle='solid')



	def rate_line(self,bus,li):
	    # Tạo các điểm đỉnh của hình chữ nhật
	    x1=0
	    y1=0
	    for bus in self.line[li]:
	    	x1+=self.coor_bus[bus][0]
	    	y1+=self.coor_bus[bus][1]

	    ## tâm của hình vuông 	
	    x2=x1/2
	    y2=y1/2

	    size1=0.1

	    x3 = [x2-size1, x2+size1, x2+size1, x2-size1,x2-size1 ]
	    y3 = [y2-size1, y2-size1, y2+size1, y2+size1, y2-size1]

	    	# Vẽ đồ thị
	    plt.fill(x3, y3,color='black')
	    plt.axis('equal')  
	    plt.grid(True)


		
def test():
	return

if __name__ == '__main__':
	fi='Inputs33bus.xlsx'
	size=10
   	
	
	option=['rate_line']
	Plot(fi).main()


	plt.show()



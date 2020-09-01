import openpyxl
import pygame
import sys

WHITE = (255,255,255)
BLACK = (0,0,0)

NAVY = (0,0,128)
ROYALBLUE = (65,105,225)
PALEGREEN = (152,251,152)
LIMEGREEN = (50,205,50)
FORESTGREEN = (34,139,34)
DARKGREEN = (0,100,0)
TAN = (210,180,140)
DARKGRAY = (169,169,169)
ORANGERED = (255,69,0)
MAP_COLORS = {
	"S":NAVY, 
	"R":ROYALBLUE, 
	"G":PALEGREEN, 
	"P":LIMEGREEN, 
	"F":FORESTGREEN, 
	"M":DARKGREEN, 
	"V":TAN, 
	"C":DARKGRAY,
	"T":ORANGERED,
	"W":TAN
}

def get_mapdata():
	wb = openpyxl.load_workbook("Map.xlsx")
	#print(wb.sheetnames)
	mapsheet = wb["Map"]

	sheetData = []
	for row in mapsheet.rows:
		sheetData.append(row)
	#print("ColNum:" + str(len(row)))
	#print("RowNum:" + str(len(sheetData)))

	# 配列化
	vals = []
	for row in sheetData:
		rowVals = []
		for cell in row:
			if not cell.value:
				rowVals.append("")
			else:
				rowVals.append(cell.value)
		vals.append(rowVals)
	#print(vals)

	# 文字表示のため、文字列で結合
	for row in sheetData:
		rowStr = ""
		for cell in row:
			if not cell.value:
				rowStr = rowStr + " "
			else:
				rowStr = rowStr + cell.value
		#print(rowStr)
	return vals

def create_map(bg, mouse_x, mouse_y, mapVals):
	bg.fill(WHITE)
	SQUARE_SIZE = 20
	
	# 画面のサイズ
	maxRectsW = -((-1 * bg.get_width()) // SQUARE_SIZE)
	maxRectsH = -((-1 * bg.get_height()) // SQUARE_SIZE)
	
	# 
	mapUpperRects = maxRectsH // 2	# 切り捨て
	mapDownerRects = -((-1 * maxRectsH) // 2)	# 切り上げ
	mapLeftRects = maxRectsW // 2	# 切り捨て
	mapRightRects = -((-1 * maxRectsW) // 2)	# 切り上げ
	
	mapArray_x = -((-1 * mouse_x) // SQUARE_SIZE)
	mapArray_y = -((-1 * mouse_y) // SQUARE_SIZE)

	if mapArray_x < mapUpperRects:
		mapArray_x = mapUpperRects
	if mapArray_x > maxRectsW - mapDownerRects:
		mapArray_x = maxRectsW - mapDownerRects
	if mapArray_y < mapLeftRects:
		mapArray_y = mapLeftRects
	if mapArray_y > maxRectsH - mapRightRects:
		mapArray_y = maxRectsH - mapRightRects

	# if mapArray_x < 12:
		# mapArray_x = 12
	# if mapArray_x > 50 - 13:
		# mapArray_x = 50 - 13
	# if mapArray_y < 12:
		# mapArray_y = 12
	# if mapArray_y > 50 - 13:
		# mapArray_y = 50 - 13
	
	
	draw_start_x = -1
	draw_start_y = -1
	for y in range(len(mapVals)):
		for x in range(len(mapVals[0])):
			if (mapArray_x - mapUpperRects) <= x and x <= (mapArray_x + mapDownerRects - 1) and (mapArray_y - mapLeftRects) <= y and y <= (mapArray_y + mapRightRects):
				if draw_start_x < 0:
					draw_start_x = x
				if draw_start_y < 0:
					draw_start_y = y
#				print("x:"+str(x)+" y:"+str(y)+" mouse_x:"+str(mapArray_x)+" mouse_y:"+str(mapArray_y))
				X = (x - draw_start_x) * SQUARE_SIZE
				Y = (y - draw_start_y) * SQUARE_SIZE
				if mapVals[y][x] == "":
					pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE, SQUARE_SIZE])
				else:
					pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, SQUARE_SIZE, SQUARE_SIZE])

gloMouseX = 250
gloMouseY = 250

def main():
	global gloMouseX, gloMouseY
	
	pygame.init()
	displayInfo = pygame.display.Info()
	screen = pygame.display.set_mode((0, 0))
	clock = pygame.time.Clock()
	former_mouseX = 0
	former_mouseY = 0

	mapVals = get_mapdata()

	while True:
		for event in pygame.event.get():
			if event.type == pygame.QUIT:
				pygame.quit()
				sys.exit()
		
		create_map(screen, former_mouseX, former_mouseY, mapVals)
			
		mouseX, mouseY = pygame.mouse.get_rel()
		mBtn1, mBtn2, mBtn3 = pygame.mouse.get_pressed()
		print("mouseX:"+str(mouseX)+" former_mouseX:"+str(former_mouseX)+" mouseY:"+str(mouseY)+" former_mouseY:"+str(former_mouseY))
		if (mouseX != former_mouseX or mouseY != former_mouseY) and mBtn1 == 1:
			mouseX = (mouseX * -1) + former_mouseX
			mouseY = (mouseY * -1) + former_mouseY
			former_mouseX = mouseX
			former_mouseY = mouseY
			# create map
			create_map(screen, mouseX, mouseY, mapVals)
		
		pygame.display.update()
		clock.tick(10)

if __name__ == "__main__":
	main()

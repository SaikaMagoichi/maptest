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
	
	# 画面のサイズから、マスの数を求める
	maxRectsW = -((-1 * bg.get_width()) // SQUARE_SIZE)
	maxRectsH = -((-1 * bg.get_height()) // SQUARE_SIZE)
	
	# マップの全マス数
	mapMaxRectW = len(mapVals[0])
	mapMaxRectH = len(mapVals)
	
	# 画面上半分のマスの数
	mapUpperRects = maxRectsH // 2	# 切り捨て
	# 画面下半分のマスの数
	mapDownerRects = -((-1 * maxRectsH) // 2)	# 切り上げ
	# 画面左半分のマスの数
	mapLeftRects = maxRectsW // 2	# 切り捨て
	# 画面右半分のマスの数
	mapRightRects = -((-1 * maxRectsW) // 2)	# 切り上げ
	
	# マウス位置をマスの位置に変換
	mapArray_x = -((-1 * mouse_x) // SQUARE_SIZE)
	mapArray_y = -((-1 * mouse_y) // SQUARE_SIZE)

	# マス目サイズ以下の部分の計算
	amariX = mouse_x % SQUARE_SIZE
	amariY = mouse_y % SQUARE_SIZE

	# マウスの位置が、これ以上スクロールできないマップの端にある場合、上下左右の表示限界位置にマウスの位置を強制的に固定する
	if mapArray_x < mapUpperRects:
		mapArray_x = mapUpperRects
		amariX = 0
	if mapArray_x > mapMaxRectW - mapDownerRects:
		mapArray_x = mapMaxRectW - mapDownerRects
		amariX = 0
	if mapArray_y < mapLeftRects:
		mapArray_y = mapLeftRects
		amariY = 0
	if mapArray_y > mapMaxRectH - mapRightRects:
		mapArray_y = mapMaxRectH - mapRightRects
		amariY = 0
	
	# if mapArray_x < 12:
		# mapArray_x = 12
	# if mapArray_x > 50 - 13:
		# mapArray_x = 50 - 13
	# if mapArray_y < 12:
		# mapArray_y = 12
	# if mapArray_y > 50 - 13:
		# mapArray_y = 50 - 13
	
	
	# 画面にマップを描画する
	draw_start_x = -1
	draw_start_y = -1
	# マップの全データ分ループする
	for y in range(len(mapVals)):
		for x in range(len(mapVals[0])):
			# マップデータが、表示エリア内の場合、描画する
			if (mapArray_x - mapUpperRects) <= x and x <= (mapArray_x + mapDownerRects) and (mapArray_y - mapLeftRects) <= y and y <= (mapArray_y + mapRightRects):
				# マップ描画開始インデックスをセットする
				if draw_start_x < 0:
					draw_start_x = x
				if draw_start_y < 0:
					draw_start_y = y
					
				# 初めて描画する場合、あまりマス目の判定をし、必要なら描画
				X = 0
				Y = 0
				squareX = 0
				squareY = 0
				if x == draw_start_x:
					X = (x - draw_start_x) * SQUARE_SIZE
					squareX = SQUARE_SIZE - amariX
				else:
					X = ((x - draw_start_x) * SQUARE_SIZE) - amariX
					squareX = SQUARE_SIZE
				if y == draw_start_y:
					y = (y - draw_start_y) * SQUARE_SIZE
					squareY = SQUARE_SIZE - amariY
				else:
					Y = ((y - draw_start_y) * SQUARE_SIZE) - amariY
					squareY = SQUARE_SIZE

				# マスを描画
				if mapVals[y][x] == "":
					pygame.draw.rect(bg, WHITE, [X, Y, squareX, squareY])
				else:
					pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, squareX, squareY])
				
				# lastSquareX = 0
				# lastSquareY = 0
				# lastSquareIndexX = x
				# lastSquareIndexY = y
				# if x == (mapArray_x + maxRectsW) and amariX > 0:
					# lastSquareX = (x - draw_start_x + 1) * SQUARE_SIZE - amariX
					# lastSquareIndexX += lastSquareIndexX
					# squareX = amariX
				# if y == (mapArray_y + maxRectsH) and amariY > 0:
					# lastSquareY = (y - draw_start_y + 1) * SQUARE_SIZE - amariY
					# lastSquareIndexY += lastSquareIndexY
					# squareY = amariY
				# if lastSquareX > 0 or lastSquareY > 0:
					# print("lastSquareIndexX:"+str(lastSquareIndexX)+" lastSquareIndexY:"+str(lastSquareIndexY))
					# if mapVals[lastSquareIndexY][lastSquareIndexX] == "":
						# pygame.draw.rect(bg, WHITE, [lastSquareX, lastSquareY, squareX, squareY])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[lastSquareIndexY][lastSquareIndexX]], [lastSquareX, lastSquareY, squareX, squareY])
				
				
				# if x == draw_start_x and amariX > 0 and y == draw_start_y and amariY > 0:
					# X = (x - draw_start_x) * SQUARE_SIZE
					# Y = (y - draw_start_y) * SQUARE_SIZE
					# if mapVals[y-1][x-1] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE - amariX, SQUARE_SIZE])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y-1][x-1]], [X, Y, SQUARE_SIZE - amariX, SQUARE_SIZE - amariY])
				# if x == draw_start_x and amariX > 0:
					# X = (x - draw_start_x) * SQUARE_SIZE
					# Y = (y - draw_start_y) * SQUARE_SIZE
					# if mapVals[y][x-1] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE - amariX, SQUARE_SIZE])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x-1]], [X, Y, SQUARE_SIZE - amariX, SQUARE_SIZE])
				# if y == draw_start_y and amariY > 0:
					# X = (x - draw_start_x) * SQUARE_SIZE
					# Y = (y - draw_start_y) * SQUARE_SIZE
					# if mapVals[y-1][x] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE, SQUARE_SIZE -amariY])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y-1][x]], [X, Y, SQUARE_SIZE, SQUARE_SIZE - amariY])
					
#				print("x:"+str(x)+" y:"+str(y)+" mouse_x:"+str(mapArray_x)+" mouse_y:"+str(mapArray_y))

					
				# # 描画座標
				# if amariX > 0:
					# X = (x - draw_start_x) * SQUARE_SIZE + SQUARE_SIZE - amariX
				# else:
					# X = (x - draw_start_x) * SQUARE_SIZE
				# if amariY > 0:
					# Y = (y - draw_start_y) * SQUARE_SIZE + SQUARE_SIZE - amariY
				# else:
					# Y = (y - draw_start_y) * SQUARE_SIZE
				
				# if mapVals[y][x] == "":
					# pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE, SQUARE_SIZE])
				# else:
					# pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, SQUARE_SIZE, SQUARE_SIZE])

				# # 右端のあまり描画
				# if amariX > 0 and x == (maxRectsW + draw_start_x) and amariY > 0 and y == (maxRectsH + draw_start_y):
					# X = (x - draw_start_x) * SQUARE_SIZE + amariX
					# Y = (y - draw_start_y) * SQUARE_SIZE + amariY
					# if mapVals[y][x] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, amariX, SQUARE_SIZE])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, amariX, amariY])
					# break
				# if amariX > 0 and x == maxRectsW + draw_start_x:
					# X = (x - draw_start_x) * SQUARE_SIZE + amariX
					# Y = (y - draw_start_y) * SQUARE_SIZE
					# if mapVals[y][x] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, amariX, SQUARE_SIZE])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, amariX, SQUARE_SIZE])
					# break
				# if amariY > 0 and y == maxRectsH + draw_start_y:
					# X = (x - draw_start_x) * SQUARE_SIZE
					# Y = (y - draw_start_y) * SQUARE_SIZE + amariY
					# if mapVals[y][x] == "":
						# pygame.draw.rect(bg, WHITE, [X, Y, SQUARE_SIZE, amariY])
					# else:
						# pygame.draw.rect(bg, MAP_COLORS[mapVals[y][x]], [X, Y, SQUARE_SIZE, amariY])
					# break

gloMouseX = 250
gloMouseY = 250

def main():
	global gloMouseX, gloMouseY
	
	pygame.init()
	displayInfo = pygame.display.Info()
#	screen = pygame.display.set_mode((0, 0))
	screen = pygame.display.set_mode((500, 500))
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

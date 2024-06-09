import os
import win32com.client
from script_api import constants as constants

people_count = int(input("Введите количество людей "))
app = win32com.client.Dispatch("Illustrator.Application")
doc = app.Documents.Add(DocumentColorSpace=1, Width=810, Height=1440)
active_doc = app.ActiveDocument

# Цвета
blue_color = win32com.client.Dispatch("Illustrator.RGBColor")
blue_color.Red = 0
blue_color.Green = 225
blue_color.Blue = 255

white_color = win32com.client.Dispatch("Illustrator.RGBColor")
white_color.Red = 255
white_color.Green = 255
white_color.Blue = 255

black_color = win32com.client.Dispatch("Illustrator.RGBColor")
black_color.Red = 0
black_color.Green = 0
black_color.Blue = 0

darkblue_color = win32com.client.Dispatch("Illustrator.RGBColor")
darkblue_color.Red = 0
darkblue_color.Green = 144
darkblue_color.Blue = 178

rain_color = win32com.client.Dispatch("Illustrator.RGBColor")
rain_color.Red = 149
rain_color.Green = 243
rain_color.Blue = 255

# Маска для лого
# -Задний фон-
layer = doc.Layers.Add()
layer.Name = "Mask_layer"
background = layer.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
background.Name = "background"
background.FillColor = rain_color
background.Stroked = False

# Группа с маской
group_mask = layer.GroupItems.Add()
group_mask.Name = "Mask"
white_background = group_mask.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
white_background.Name = "rain_color-background"
white_background.FillColor = white_color
white_background.Stroked = False

ellipse_mask = group_mask.PathItems.Ellipse(692, 512, 147, 147)
ellipse_mask.Name = "Circle"
ellipse_mask.Filled = False
ellipse_mask.Stroked = False

group_mask.Clipped = True

# Главный фон
layer1 = doc.Layers.Add()
layer1.Name = "BG"
BG_path_group = layer1.GroupItems.Add()
BG_path_group.Name = "Path_group"
BG_rectangle = BG_path_group.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
BG_rectangle.Name = "BG_rectangle"
BG_rectangle.FillColor = blue_color
BG_rectangle.Stroked = False

# -Дождь-
rain_group = BG_path_group.GroupItems.Add()
rain_group.Name = "Rain"
rain1 = rain_group.PathItems.RoundedRectangle(800, 250, 85, 10, 100, 100)
rain1.Rotate(45)
rain1.Name = "rain1"
rain1.FillColor = white_color
rain1.Stroked = False

rain2 = rain_group.PathItems.RoundedRectangle(692, 410, 50, 15, 100, 100)
rain2.Rotate(45)
rain2.Name = "rain2"
rain2.FillColor = white_color
rain2.Stroked = False

rain3 = rain_group.PathItems.RoundedRectangle(830, 510, 20, 85, 100, 100)
rain3.Rotate(-45)
rain3.Name = "rain3"
rain3.FillColor = white_color
rain3.Stroked = False

rain4 = rain_group.PathItems.RoundedRectangle(710, 270, 65, 10, 100, 100)
rain4.Rotate(45)
rain4.Name = "rain4"
rain4.FillColor = white_color
rain4.Stroked = False

rain5 = rain_group.PathItems.RoundedRectangle(685, 220, 20, 7, 100, 100)
rain5.Rotate(45)
rain5.Name = "rain5"
rain5.FillColor = white_color
rain5.Stroked = False

rain6 = rain_group.PathItems.RoundedRectangle(630, 195, 35, 10, 100, 100)
rain6.Rotate(45)
rain6.Name = "rain6"
rain6.FillColor = white_color
rain6.Stroked = False

rain7 = rain_group.PathItems.RoundedRectangle(730, 395, 45, 10, 100, 100)
rain7.Rotate(45)
rain7.Name = "rain7"
rain7.FillColor = white_color
rain7.Stroked = False

rain8 = rain_group.PathItems.RoundedRectangle(630, 415, 55, 9, 100, 100)
rain8.Rotate(45)
rain8.Name = "rain8"
rain8.FillColor = white_color
rain8.Stroked = False

rain9 = rain_group.PathItems.RoundedRectangle(780, 455, 20, 5, 100, 100)
rain9.Rotate(45)
rain9.Name = "rain9"
rain9.FillColor = white_color
rain9.Stroked = False

rain10 = rain_group.PathItems.RoundedRectangle(645, 320, 45, 7, 100, 100)
rain10.Rotate(45)
rain10.Name = "rain10"
rain10.FillColor = white_color
rain10.Stroked = False

rain11 = rain_group.PathItems.RoundedRectangle(790, 565, 50, 85, 100, 100)
rain11.Rotate(-45)
rain11.Name = "rain11"
rain11.FillColor = white_color
rain11.Stroked = False

rain12 = rain_group.PathItems.RoundedRectangle(660, 285, 30, 8, 100, 100)
rain12.Rotate(45)
rain12.Name = "rain12"
rain12.FillColor = white_color
rain12.Stroked = False

# -Лого-
circle = BG_path_group.PathItems.Ellipse(692, 512, 147, 147)
circle.Name = "Circle"
circle.FillColor = white_color
circle.Stroked = False

# --Иконка лого--
layer_logo = doc.Layers.Add()
layer_logo.Name = "LogoLayer"
logo_icon_file = "media/logo_icon.png"
logo_icon_file_path = os.path.abspath(logo_icon_file)
placedItem = doc.PlacedItems.Add()
placedItem.File = logo_icon_file_path
placedItem.Name = "Logo_icon"
placedItem.Move(layer_logo, 1)
placedItem.Width = 44
placedItem.Height = 44
placedItem.Left = 533
placedItem.Top = 670

# Внутренняя плашка
layer2 = doc.Layers.Add()
layer2.Name = "Title"

# -Человечки-
if people_count > 0:
    layer2_1 = layer2.GroupItems.Add()
    layer2_1.Name = "Human"
    human_file = "media/human.png"
    human_path = os.path.abspath(human_file)
    placedItem = doc.PlacedItems.Add()
    placedItem.File = human_path
    placedItem.Name = "Man"
    placedItem.Move(layer2_1, 1)
    placedItem.Width = 10
    placedItem.Height = 28
    placedItem.Left = 279
    placedItem.Top = 768

    for i in range(people_count-1):
        placedItem.Duplicate(layer2_1, constants.aiPlaceInside)
        placedItem.Left += 14
else:
    pass

# -Фон внутренней плашки-
title_rectangle = layer2.PathItems.RoundedRectangle(755, 230, 220, 36, 10, 10)
title_rectangle.Name = "Title_rectangle"
title_rectangle.FillColor = darkblue_color
title_rectangle.Stroked = False

# -Тень текста-
group = layer2.GroupItems.Add()
group.Name = "Text_shadow"
text = group.TextFrames.PointText(title_rectangle.Position)
text.Contents = "КАТЕГОРИЯ"
category = text.Contents
text.Position = [text.Position[0] + 108, text.Position[1] - 27]
text_range = text.TextRange
text_range.ParagraphAttributes.Justification = 2
text_range.CharacterAttributes.Size = 22
text_range.CharacterAttributes.FillColor = black_color
text_range.CharacterAttributes.TextFont = app.TextFonts.Item("YuGothicUI-Semibold")
group.Opacity = 26

# -Текст-
text2 = layer2.TextFrames.PointText(title_rectangle.Position)
text2.Contents = category
text2.Position = [text.Position[0] + 63, text.Position[1] - 8]
text_range2 = text2.TextRange
text_range2.ParagraphAttributes.Justification = 2
text_range2.CharacterAttributes.Size = 22
text_range2.CharacterAttributes.FillColor = white_color
text_range2.CharacterAttributes.TextFont = app.TextFonts.Item("YuGothicUI-Semibold")

# Тень внутренней плашки
layer3 = doc.Layers.Add()
layer3.Name = "Title Inner Shadow"

title_shadow = layer3.PathItems.RoundedRectangle(755, 230, 220, 36, 10, 10)
title_shadow.Name = "Black Fill"
title_shadow.FillColor = black_color
title_shadow.Stroked = False

title_shadow2 = layer3.PathItems.RoundedRectangle(755, 230, 220, 36, 10, 10)
title_shadow2.Name = "White Fill"
title_shadow2.FillColor = white_color
title_shadow2.Stroked = False
title_shadow2.ApplyEffect('<LiveEffect name="Adobe PSL Gaussian Blur"><Dict data="R blur 10 "/></LiveEffect>')
title_shadow2.ApplyEffect('<LiveEffect name="Adobe PSL Transform"><Dict data="B transformPatterns 1 R moveH_Pts 0 R moveV_Pts -3 B transformObjects 1" /></LiveEffect>')

layer3.BlendingMode = 1
layer3.Opacity = 68

# Тень главного фона
layer4 = doc.Layers.Add()
layer4.Name = "BG Inner Shadow"

shadow_rectangle = layer4.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
shadow_rectangle.Name = "Black Fill"
shadow_rectangle.FillColor = black_color
shadow_rectangle.Stroked = False

shadow_rectangle2 = layer4.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
shadow_rectangle2.Name = "White Fill"
shadow_rectangle2.FillColor = white_color
shadow_rectangle2.Stroked = False
shadow_rectangle2.ApplyEffect('<LiveEffect name="Adobe PSL Gaussian Blur"><Dict data="R blur 10 "/></LiveEffect>')
shadow_rectangle2.ApplyEffect('<LiveEffect name="Adobe PSL Transform"><Dict data="B transformPatterns 1 R moveH_Pts 4 R moveV_Pts -4 B transformObjects 1" /></LiveEffect>')

layer4.BlendingMode = 1
layer4.Opacity = 70

# Обрезка фигур, выходящих за рамки
BG_path_group.Selected = True
app.ExecuteMenuCommand('Live Pathfinder Subtract')

# Обрезка эффектов, выходящих за рамки
# layer5 = doc.Layers.Add()
# layer5.Name = "Outer Cut"
# group_mask2 = layer5.GroupItems.Add()
# group_mask2.Name = "Mask2"
# outer_cut_rectangle = group_mask2.PathItems.RoundedRectangle(800, 200, 385, 182, 10, 10)
# outer_cut_rectangle.Filled = False
# outer_cut_rectangle.Stroked = False


# Сохранение
output_file = os.path.abspath("test/test")
save_options = constants.aiPNG24
active_doc.Export(output_file, save_options)
doc.Close(2)
app.Quit()

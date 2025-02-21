from PIL import Image, ImageDraw, ImageFont
import fitz
from io import BytesIO
import numpy as np
import pandas as pd
from Path import path
import os 
from Utils import utils

class main_treatment:

    def stamp (self, op):

        df_planner = pd.read_excel(path.path_documents + "Documents.xlsx", sheet_name='Planners')
        df_planner = df_planner[df_planner['User'].astype(str) == str(os.getcwd().split(os.sep)[2])]

        if df_planner.empty:
            return "User not found"
        
        df_planner = df_planner.reset_index(drop=True)
        name = df_planner.at[0,'Name']
        id = df_planner.at[0,'ID']
            
        #Definir as informações do carimbo
        title = "LIBERADO PARA EXECUÇÃO"
        red = (255, 0, 0)  #Vermelho
        white = (255, 255, 255)  #Branco
        stamp_width = 300
        stamp_heigth = 200
        stamp_border = 5
        title_font = ImageFont.truetype("arial.ttf", 20)  #Fonte maior para o título
        text_font = ImageFont.truetype("arial.ttf", 18)  #Fonte menor para o restante

        #Criar imagem base
        image = Image.new("RGB", (stamp_width, stamp_heigth), white)
        draw = ImageDraw.Draw(image)

            #Desenhar borda do carimbo
        draw.rectangle(
            [(stamp_border, stamp_border), (stamp_width-stamp_border, stamp_heigth-stamp_border)],
            outline=red,
            width=stamp_border
        )

        title_bbox = draw.textbbox((0, 0), title, font=title_font)  # Retorna (x0, y0, x1, y1)
        w_title = title_bbox[2] - title_bbox[0]  # x1 - x0
        draw.text(((stamp_width - w_title) / 2, 30), title, font=title_font, fill=red)
            
        op_bbox = draw.textbbox((0, 0), f"OP {op}", font=title_font)  # Retorna (x0, y0, x1, y1)
        w_op = op_bbox[2] - op_bbox[0]  # x1 - x0
        draw.text(((stamp_width - w_op) / 2, 70), f"OP {op}", font=title_font, fill=red)
        
        author_bbox = draw.textbbox((0, 0), f"{name} - REG {id}", font=text_font)  # Retorna (x0, y0, x1, y1)
        w_author = author_bbox[2] - author_bbox[0]  # x1 - x0
        draw.text(((stamp_width - w_author) / 2, 110), f"{name} - REG {id}", font=text_font, fill=red)
        
        data_bbox = draw.textbbox((0, 0), f'Printed on {utils.get_time()}', font=text_font)  # Retorna (x0, y0, x1, y1)
        w_data = data_bbox[2] - data_bbox[0]  # x1 - x0
        draw.text(((stamp_width - w_data) / 2, 150), f'Printed on {utils.get_time()}', font=text_font, fill=red)
        
        # Salvar a imagem em um buffer de memória BytesIO
        img_byte_arr = BytesIO()
        image.save(img_byte_arr, format="PNG")
        img_byte_arr.seek(0)  # Voltar o cursor para o início

         # Criar um Pixmap a partir dos bytes da imagem
        pixmap = fitz.Pixmap(img_byte_arr)

        #Salvar ou mostrar a imagem
        return pixmap #retornar como imagem
    
    def find_space(self, page, image_width, image_height, margin=10,threshold=240):
        
        # Obter a imagem da página
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Converter a imagem para escala de cinza
        grayscale_img = img.convert("L")
        img_array = np.array(grayscale_img)

        # Percorrer a página para encontrar uma área em branco
        for x in range(0, pix.width - image_width, margin):  # Passos de 10 pixels para eficiência
            for y in range(0, pix.height - image_height, margin):
                region = img_array[y:y + image_height, x:x + image_width]
                
                # Verificar a média da região - se for mais clara que o limiar, consideramos "branco"
                avg_brightness = np.mean(region)
                
                # Se a média de brilho for maior que o limiar, é uma área potencialmente em branco
                if avg_brightness > threshold:
                    # Verificar se a região não contém textos ou linhas (apenas áreas muito claras)
                    if np.all(region > threshold):  # Garantir que todos os pixels da região são claros
                        return x, y   
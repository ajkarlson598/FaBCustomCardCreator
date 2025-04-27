import pandas as pd
from PIL import ImageFont, ImageDraw, Image
import textwrap

def wrappingEffects (draw: ImageDraw, font: ImageFont ,input, xCAT, yCAT):
    lines = textwrap.wrap(input, width=27)
    for line in lines:
        draw.text((xCAT, yCAT), line, fill='black', font=font)
        yCAT += 20
    return yCAT

def readingExcel():
    return pd.read_excel('fabpve.xlsx', sheet_name='entity_cards')

def drawCardBoarder(draw: ImageDraw, TLV: int, TRV: int, BLV: int, BRV: int):
    draw.line([TLV,TRV], fill='black', width=6)
    draw.line([TRV,BRV], fill='black', width=6)
    draw.line([BRV,BLV], fill='black', width=6)
    draw.line([BLV,TLV], fill='black', width=6)
    return draw

def populateCardText(draw: ImageDraw, row: pd.Series , font: ImageFont, TLV: int):
    draw.text((TLV[0] + 60, TLV[1] + 10), f"{row['Name']}", fill='black', font=font)
    draw.text((TLV[0] + 75, TLV[1] + 240 ), f"{row['Card Type']}", fill='black', font=font)
    draw.text((TLV[0] +  20, TLV[1] + 10), f"{row['Cost']}", fill='Red', font=font)
    draw.text((TLV[0] + 170, TLV[1] + 10), f"{row['Pitch']}", fill='black', font=font)
    draw.text((TLV[0] + 20, TLV[1] + 240), f"{row['Power']}", fill='Gold', font=font)
    draw.text((TLV[0] + 170, TLV[1] + 240), f"{row['Defense']}", fill='Grey', font=font)
    effectsYValue = wrappingEffects(draw=draw, font=font ,input=f"{row['Effect 1']}", xCAT=TLV[0] + 10, yCAT=TLV[1] + 50)
    effectsYValue = wrappingEffects(draw=draw, font=font ,input=f"{row['Effect 2']}", xCAT=TLV[0]+10, yCAT=effectsYValue + 10)
    effectsYValue = wrappingEffects(draw=draw, font=font ,input=f"{row['Effect 3']}", xCAT=TLV[0]+10, yCAT=effectsYValue + 10)
    return draw

def main():
    dataframe=readingExcel()
    font = ImageFont.load_default(15)

    card = Image.new('RGB', (1100, 850), 'white')
    draw = ImageDraw.Draw(card)

    topleftVal = (0,5)
    topRightVal = (200,5)
    bottomLeftVal = (0,280)
    bottomRightVal = (200,280)

    for index, row in dataframe.iterrows():
        if(index%5 == 0 and index !=0):
            # newline for the cards
            modifier = 280 * (index/5)
            topleftVal = (0, 5 + modifier)
            topRightVal = (200, 5 + modifier)
            bottomLeftVal = (0, 280 + modifier)
            bottomRightVal = (200,280 + modifier)
        if((index)%15==0 and index!=0):
            filename = f"output_cards/{row['Name'].replace(' ', '_')}.png"
            card.save(filename, dpi=(300,300))
            print(f"Saved: {filename}")

            card = Image.new('RGB', (1100, 850), 'white')
            draw = ImageDraw.Draw(card)
            topleftVal = (0,5)
            topRightVal = (200,5)
            bottomLeftVal = (0,280)
            bottomRightVal = (200,280)
    
        draw = drawCardBoarder(draw, TLV=topleftVal, TRV=topRightVal, BLV=bottomLeftVal, BRV=bottomRightVal) 
        draw = populateCardText(draw=draw, row=row, font=font, TLV=topleftVal)

        topleftVal = topRightVal
        topRightVal = (topRightVal[0] + 200, topRightVal[1])
        bottomLeftVal = bottomRightVal
        bottomRightVal = (bottomRightVal[0] + 200, bottomRightVal[1])

    filename = f"output_cards/{row['Name'].replace(' ', '_')}.png"
    card.save(filename, dpi=(300,300))
    print(f"Saved: {filename}")

if __name__=="__main__":
    main()




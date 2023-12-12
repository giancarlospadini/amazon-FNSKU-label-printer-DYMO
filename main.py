import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import font
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# check driver status:
# https://localhost:41951/DYMO/DLS/Printing/StatusConnected
# https://127.0.0.1:41951/DYMO/DLS/Printing/StatusConnected

# list connected printers:
# https://localhost:41951/DYMO/DLS/Printing/GetPrinters
# https://127.0.0.1:41951/DYMO/DLS/Printing/GetPrinters


def print_labels(barcode, text, quantity):
    
    label_xml = f"""<?xml version="1.0" encoding="utf-8"?>
    <DesktopLabel Version="1">
      <DYMOLabel Version="3">
        <Description>DYMO Label</Description>
        <Orientation>Portrait</Orientation>
        <LabelName>S0722540 multipurpose</LabelName>
        <InitialLength>0</InitialLength>
        <BorderStyle>SolidLine</BorderStyle>
        <DYMORect>
          <DYMOPoint>
            <X>0.03999997</X>
            <Y>0.06</Y>
          </DYMOPoint>
          <Size>
            <Width>2.17</Width>
            <Height>1.13</Height>
          </Size>
        </DYMORect>
        <BorderColor>
          <SolidColorBrush>
            <Color A="1" R="0" G="0" B="0"></Color>
          </SolidColorBrush>
        </BorderColor>
        <BorderThickness>1</BorderThickness>
        <Show_Border>False</Show_Border>
        <DynamicLayoutManager>
          <RotationBehavior>ClearObjects</RotationBehavior>
          <LabelObjects>
            <TextObject>
              <Name>Testo</Name>
              <Brushes>
                <BackgroundBrush>
                  <SolidColorBrush>
                    <Color A="0" R="1" G="1" B="1"></Color>
                  </SolidColorBrush>
                </BackgroundBrush>
                <BorderBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </BorderBrush>
                <StrokeBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </StrokeBrush>
                <FillBrush>
                  <SolidColorBrush>
                    <Color A="0" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </FillBrush>
              </Brushes>
              <Rotation>Rotation0</Rotation>
              <OutlineThickness>1</OutlineThickness>
              <IsOutlined>False</IsOutlined>
              <BorderStyle>SolidLine</BorderStyle>
              <Margin>
                <DYMOThickness Left="0" Top="0" Right="0" Bottom="0" />
              </Margin>
              <HorizontalAlignment>Center</HorizontalAlignment>
              <VerticalAlignment>Middle</VerticalAlignment>
              <FitMode>None</FitMode>
              <IsVertical>False</IsVertical>
              <FormattedText>
                <FitMode>None</FitMode>
                <HorizontalAlignment>Center</HorizontalAlignment>
                <VerticalAlignment>Middle</VerticalAlignment>
                <IsVertical>False</IsVertical>
                <LineTextSpan>
                  <TextSpan>
                    <Text>{text}</Text>
                    <FontInfo>
                      <FontName>Arial</FontName>
                      <FontSize>10</FontSize>
                      <IsBold>False</IsBold>
                      <IsItalic>False</IsItalic>
                      <IsUnderline>False</IsUnderline>
                      <FontBrush>
                        <SolidColorBrush>
                          <Color A="1" R="0" G="0" B="0"></Color>
                        </SolidColorBrush>
                      </FontBrush>
                    </FontInfo>
                  </TextSpan>
                </LineTextSpan>
              </FormattedText>
              <ObjectLayout>
                <DYMOPoint>
                  <X>0.05999997</X>
                  <Y>0.08</Y>
                </DYMOPoint>
                <Size>
                  <Width>2.15</Width>
                  <Height>0.6402779</Height>
                </Size>
              </ObjectLayout>
            </TextObject>
            <BarcodeObject>
              <Name>BarcodeObject0</Name>
              <Brushes>
                <BackgroundBrush>
                  <SolidColorBrush>
                    <Color A="1" R="1" G="1" B="1"></Color>
                  </SolidColorBrush>
                </BackgroundBrush>
                <BorderBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </BorderBrush>
                <StrokeBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </StrokeBrush>
                <FillBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </FillBrush>
              </Brushes>
              <Rotation>Rotation0</Rotation>
              <OutlineThickness>1</OutlineThickness>
              <IsOutlined>False</IsOutlined>
              <BorderStyle>SolidLine</BorderStyle>
              <Margin>
                <DYMOThickness Left="0" Top="0" Right="0" Bottom="0" />
              </Margin>
              <BarcodeFormat>Code128Auto</BarcodeFormat>
              <Data>
                <MultiDataString>
                  <DataString></DataString>
                  <DataString>{barcode}</DataString>
                </MultiDataString>
              </Data>
              <HorizontalAlignment>Center</HorizontalAlignment>
              <VerticalAlignment>Middle</VerticalAlignment>
              <Size>SmallMedium</Size>
              <TextPosition>Bottom</TextPosition>
              <FontInfo>
                <FontName>Arial</FontName>
                <FontSize>12</FontSize>
                <IsBold>False</IsBold>
                <IsItalic>False</IsItalic>
                <IsUnderline>False</IsUnderline>
                <FontBrush>
                  <SolidColorBrush>
                    <Color A="1" R="0" G="0" B="0"></Color>
                  </SolidColorBrush>
                </FontBrush>
              </FontInfo>
              <ObjectLayout>
                <DYMOPoint>
                  <X>0.05999997</X>
                  <Y>0.6250001</Y>
                </DYMOPoint>
                <Size>
                  <Width>2.13</Width>
                  <Height>0.5476041</Height>
                </Size>
              </ObjectLayout>
            </BarcodeObject>
          </LabelObjects>
        </DynamicLayoutManager>
      </DYMOLabel>
      <LabelApplication>Blank</LabelApplication>
      <DataTable>
        <Columns></Columns>
        <Rows></Rows>
      </DataTable>
    </DesktopLabel>"""

    label_xml = label_xml.strip().replace('\n', ''  )

    url = "https://127.0.0.1:41951/DYMO/DLS/Printing/PrintLabel"
    header = {'Content-Type': 'application/x-www-form-urlencoded'}
    body2 = {
                "printerName": "Dymo LabelWriter 450",  
                "labelXml": label_xml
    }

    body = [k + '=' + v for k, v in body2.items()]
    complete_body = "&".join(body)

    i = 1
    while i <= quantity:
      requests.post(
          url,
          headers=header,
          data=complete_body,
          verify=False
      )
      i += 1


# Funzione principale
def main():
    excel_file = excel_file_entry.get()
    if excel_file:
        try:
            wb = openpyxl.load_workbook(excel_file)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=6, values_only=True):
                barcode, text, quantity = row[4], row[0], row[9] # Considera solo le colonne interessate
                if (quantity != None):
                  print_labels(barcode, text, quantity)
            status_label.config(text="Etichette stampate con successo.")
        except Exception as e:
            status_label.config(text=f"Errore durante la lettura di Excel: {str(e)}")
    else:
        status_label.config(text="Seleziona un file Excel.")

# Funzione per selezionare il file Excel
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_file_entry.delete(0, tk.END)
        excel_file_entry.insert(0, file_path)

# Configura la root Tkinter
root = tk.Tk()
root.title("Stampa Etichette da Excel")
root.geometry("400x200")

frame = tk.Frame(root)
frame.pack(pady=20)

excel_file_label = tk.Label(frame, text="Seleziona il file Excel:")
excel_file_label.grid(row=0, column=0)

excel_file_entry = tk.Entry(frame, width=30)
excel_file_entry.grid(row=0, column=1)

browse_button = tk.Button(frame, text="Sfoglia", command=browse_file)
browse_button.grid(row=0, column=2)

print_button = tk.Button(frame, text="Stampa Etichette", command=main)
print_button.grid(row=1, column=1)

status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()
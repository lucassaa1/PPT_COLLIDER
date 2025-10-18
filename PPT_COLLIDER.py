import os
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, asksaveasfilename
import win32com.client

def unir_presentaciones():
    # Inicializar Tkinter para seleccionar archivos
    Tk().withdraw()

    # Seleccionar las presentaciones a unir
    archivos = askopenfilenames(title="Selecciona los archivos PPTX a unir", filetypes=[("PowerPoint files", "*.pptx")])
    if len(archivos) < 2:
        print("Debes seleccionar al menos 2 archivos.")
        return

    # Seleccionar dónde guardar la presentación final
    archivo_salida = asksaveasfilename(title="Guardar presentación combinada como", defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
    if not archivo_salida:
        print("No se seleccionó archivo de salida.")
        return

    # Iniciar PowerPoint
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    try:
        # Crear presentación vacía
        pres_final = ppt_app.Presentations.Add()

        # Iterar sobre cada presentación seleccionada
        for archivo in archivos:
            pres_temp = ppt_app.Presentations.Open(os.path.abspath(archivo), WithWindow=False)
            for slide in pres_temp.Slides:
                slide.Copy()
                pres_final.Slides.Paste()
            pres_temp.Close()

        # Guardar presentación final
        pres_final.SaveAs(os.path.abspath(archivo_salida))
        pres_final.Close()
        print(f"Presentaciones combinadas correctamente en: {archivo_salida}")

    except Exception as e:
        print("Ocurrió un error:", e)

    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    unir_presentaciones()

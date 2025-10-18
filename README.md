# PPT_COLLIDER
Unir múltiples presentaciones de PowerPoint (.pptx) en un solo archivo de manera rápida, confiable y profesional, conservando todos los elementos intactos.

# Características principales

Combina cualquier cantidad de presentaciones seleccionadas por el usuario.

Mantiene 100% de fidelidad: textos, imágenes, gráficos, tablas, SmartArt, animaciones, transiciones y fondos.

Funciona mediante PowerPoint real a través de COM (win32com.client), evitando problemas de corrupción de archivos.

Permite elegir fácilmente las presentaciones de origen y la ubicación del archivo final mediante un diálogo gráfico.

Fácil de usar, robusto y confiable para entornos profesionales.

Requisitos

Windows con Microsoft PowerPoint instalado.

Python 3.6+

Librerías Python:

pywin32

tkinter (generalmente incluido por defecto en Python)

Instalación de la librería necesaria:

pip install pywin32

Uso

Ejecuta el script ppt_collider.py.

Selecciona los archivos .pptx que deseas unir.

Elige la ubicación y nombre del archivo final.

¡Tu presentación combinada se generará automáticamente, manteniendo todo el contenido intacto!

Nota: Este proyecto está diseñado para entornos Windows con PowerPoint instalado. No es compatible con Linux o macOS debido al uso de COM para controlar PowerPoint.

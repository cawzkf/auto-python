import webbrowser
import os

pagina_web = ("D:\OneDrive\Documentos\Camila\Github\Html e css\modelo_web\index.html")

opera = r"C:\Users\Marcus\AppData\Local\Programs\Opera GX\opera.exe"

webbrowser.register('opera',None,webbrowser.BackgroundBrowser(opera))

webbrowser.get('opera').open(f"file://{pagina_web}")
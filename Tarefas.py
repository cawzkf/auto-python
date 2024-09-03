import webbrowser as plan
import os

pagina_web = ("D:\OneDrive\Documentos\Camila\Github\Html e css\modelo_web\index.html")

opera = r"C:\Users\Marcus\AppData\Local\Programs\Opera GX\opera.exe"

plan.register('opera',None,plan.BackgroundBrowser(opera))

plan.get('opera').open(f"file://{pagina_web}")
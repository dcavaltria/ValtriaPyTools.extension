#! python3
# -*- coding: utf-8 -*-
"""
Claude AI Assistant para Revit - Compatible con IronPython 2.7
Consulta inteligente sobre elementos del modelo
"""
__title__ = "Consultar\nClaude"
__author__ = "David Carreras / Valtria"
__doc__ = "Asistente IA para anÃ¡lisis de elementos de Revit"
__context__ = 'zero-doc'

from pyrevit import revit, DB, forms, script
from System.Windows import Window
from System.Windows.Controls import TextBox, Button, StackPanel, ScrollViewer, TextBlock
import wpf
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import json
import os
import sys
from datetime import datetime
import hashlib

SCRIPT_DIR = os.path.dirname(__file__)

# Agregar rutas adicionales para buscar anthropic (sin permisos de admin)
user_site = os.path.expanduser(r'~\.pyrevit\site-packages')
pyrevit_site = r'C:\Program Files\pyRevit-Master\site-packages'

for path in [user_site, pyrevit_site]:
    if os.path.exists(path) and path not in sys.path:
        sys.path.insert(0, path)

try:
    import anthropic
except ImportError as imp_err:
    try:
        log_path = os.path.join(SCRIPT_DIR, "anthropic_import_error.log")
        with open(log_path, 'w') as log_file:
            log_file.write("ImportError: {}\n".format(imp_err))
            import traceback
            traceback.print_exc(file=log_file)
    except Exception:
        pass
    forms.alert(
        u"âŒ Falta la librerÃ­a 'anthropic'\n\n"
        u"SIN PERMISOS DE ADMIN, instala en tu carpeta de usuario:\n\n"
        u"python -m pip install anthropic --target=\"{}\"  \n\n"
        u"O ejecuta el archivo en tu Escritorio:\n"
        u"instalar_anthropic_SIN_ADMIN.bat\n\n"
        u"O descarga manualmente desde:\n"
        u"https://pypi.org/project/anthropic/#files".format(user_site),
        exitscript=True
    )

# ============= CONFIGURACIÃ“N =============
API_CONFIG_FILE = os.path.join(SCRIPT_DIR, "api_config.json")
CONTADOR_FILE = os.path.join(SCRIPT_DIR, "contador_consultas.json")
CACHE_FILE = os.path.join(SCRIPT_DIR, "cache_respuestas.json")

def cargar_configuracion():
    """Carga la configuraciÃ³n desde el archivo JSON"""
    try:
        if os.path.exists(API_CONFIG_FILE):
            with open(API_CONFIG_FILE, 'r') as f:
                config = json.load(f)
                return config
        else:
            config_default = {
                "api_key": "TU_API_KEY_AQUI",
                "usar_haiku": True,
                "max_consultas_dia": 50
            }
            with open(API_CONFIG_FILE, 'w') as f:
                json.dump(config_default, f, indent=4)
            return config_default
    except Exception as e:
        forms.alert(u"Error al cargar configuraciÃ³n: {}".format(str(e)), exitscript=True)
        return None

CONFIG = cargar_configuracion()
API_KEY = CONFIG.get("api_key", "TU_API_KEY_AQUI")
USAR_HAIKU = CONFIG.get("usar_haiku", True)
MAX_CONSULTAS_DIA = CONFIG.get("max_consultas_dia", 50)

def cargar_contador():
    """Carga el contador de consultas del dÃ­a"""
    try:
        if os.path.exists(CONTADOR_FILE):
            with open(CONTADOR_FILE, 'r') as f:
                data = json.load(f)
                if data.get('fecha') == datetime.now().strftime('%Y-%m-%d'):
                    return data.get('contador', 0)
        return 0
    except:
        return 0

def guardar_contador(contador):
    """Guarda el contador de consultas"""
    try:
        data = {
            'fecha': datetime.now().strftime('%Y-%m-%d'),
            'contador': contador
        }
        with open(CONTADOR_FILE, 'w') as f:
            json.dump(data, f, indent=2)
    except:
        pass

def cargar_cache():
    """Carga el cachÃ© de respuestas"""
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r') as f:
                cache = json.load(f)
                fecha_actual = datetime.now().strftime('%Y-%m-%d')
                cache_limpio = {}
                for k, v in cache.items():
                    if v.get('fecha') == fecha_actual:
                        cache_limpio[k] = v
                return cache_limpio
        return {}
    except:
        return {}

def guardar_cache(cache):
    """Guarda el cachÃ© de respuestas"""
    try:
        with open(CACHE_FILE, 'w') as f:
            json.dump(cache, f, indent=2)
    except:
        pass

def generar_hash_consulta(pregunta, elementos):
    """Genera un hash Ãºnico para la consulta"""
    ids_elementos = sorted([elem['ID'] for elem in elementos])
    consulta_str = pregunta.lower().strip() + '|' + '|'.join(ids_elementos)
    return hashlib.md5(consulta_str.encode('utf-8')).hexdigest()

class ClaudeWindow(Window):
    """Ventana principal de la interfaz"""

    def __init__(self):
        self.Title = "Claude AI - Asistente Revit"
        self.Width = 600
        self.Height = 700
        self.WindowStartupLocation = 0

        self.contador_consultas = cargar_contador()
        self.cache = cargar_cache()

        panel = StackPanel()
        panel.Margin = wpf.Thickness(15)

        titulo = TextBlock()
        titulo.Text = u"ðŸ¤– Claude AI Assistant"
        titulo.FontSize = 20
        titulo.FontWeight = wpf.FontWeights.Bold
        titulo.Margin = wpf.Thickness(0, 0, 0, 15)
        panel.Children.Add(titulo)

        instrucciones = TextBlock()
        instrucciones.Text = (
            u"Selecciona elementos en Revit y haz tu pregunta.\n"
            u"Ejemplos: 'Â¿Cumplen normativa?', 'Resume parÃ¡metros', 'Encuentra errores'"
        )
        instrucciones.TextWrapping = 1
        instrucciones.Margin = wpf.Thickness(0, 0, 0, 10)
        instrucciones.Foreground = wpf.Media.Brushes.Gray
        panel.Children.Add(instrucciones)

        pregunta_label = TextBlock()
        pregunta_label.Text = u"Tu pregunta:"
        pregunta_label.FontWeight = wpf.FontWeights.Bold
        pregunta_label.Margin = wpf.Thickness(0, 10, 0, 5)
        panel.Children.Add(pregunta_label)

        self.pregunta_box = TextBox()
        self.pregunta_box.Height = 80
        self.pregunta_box.TextWrapping = 1
        self.pregunta_box.AcceptsReturn = True
        self.pregunta_box.VerticalScrollBarVisibility = 1
        panel.Children.Add(self.pregunta_box)

        btn_consultar = Button()
        btn_consultar.Content = u"ðŸš€ Consultar a Claude"
        btn_consultar.Height = 40
        btn_consultar.Margin = wpf.Thickness(0, 15, 0, 0)
        btn_consultar.FontSize = 14
        btn_consultar.Click += self.on_consultar
        panel.Children.Add(btn_consultar)

        self.info_elementos = TextBlock()
        self.info_elementos.Margin = wpf.Thickness(0, 10, 0, 10)
        self.info_elementos.Foreground = wpf.Media.Brushes.DarkBlue
        panel.Children.Add(self.info_elementos)

        respuesta_label = TextBlock()
        respuesta_label.Text = u"Respuesta de Claude:"
        respuesta_label.FontWeight = wpf.FontWeights.Bold
        respuesta_label.Margin = wpf.Thickness(0, 10, 0, 5)
        panel.Children.Add(respuesta_label)

        scroll = ScrollViewer()
        scroll.Height = 300
        scroll.VerticalScrollBarVisibility = 1

        self.respuesta_box = TextBox()
        self.respuesta_box.TextWrapping = 1
        self.respuesta_box.IsReadOnly = True
        self.respuesta_box.Background = wpf.Media.Brushes.WhiteSmoke
        scroll.Content = self.respuesta_box
        panel.Children.Add(scroll)

        self.footer = TextBlock()
        self.footer.Margin = wpf.Thickness(0, 10, 0, 0)
        self.footer.FontSize = 10
        self.footer.Foreground = wpf.Media.Brushes.Gray
        self.actualizar_footer()
        panel.Children.Add(self.footer)

        self.Content = panel

    def actualizar_footer(self):
        modelo = u"Haiku (econÃ³mico)" if USAR_HAIKU else u"Sonnet 4.5 (avanzado)"
        costo = self.contador_consultas * (0.002 if USAR_HAIKU else 0.012)
        self.footer.Text = u"Modelo: {} | Consultas hoy: {}/{} | Costo aprox: ${:.3f}".format(
            modelo, self.contador_consultas, MAX_CONSULTAS_DIA, costo
        )

    def obtener_elementos_seleccionados(self):
        """Extrae info de elementos seleccionados"""
        doc = revit.doc
        seleccion = revit.get_selection()

        if not seleccion:
            return None

        elementos = []
        for elem_id in seleccion.element_ids[:20]:
            elem = doc.GetElement(elem_id)

            if elem is None:
                continue

            info = {
                "ID": str(elem.Id),
                u"CategorÃ­a": elem.Category.Name if elem.Category else u"Sin categorÃ­a",
                "Tipo": elem.Name if hasattr(elem, 'Name') else "N/A"
            }

            params = {}
            for param in elem.Parameters:
                if param.HasValue and param.StorageType != DB.StorageType.ElementId:
                    try:
                        nombre = param.Definition.Name
                        valor = None

                        if param.AsValueString():
                            valor = param.AsValueString()
                        elif param.StorageType == DB.StorageType.Double:
                            valor = str(param.AsDouble())
                        elif param.StorageType == DB.StorageType.Integer:
                            valor = str(param.AsInteger())
                        elif param.StorageType == DB.StorageType.String:
                            valor = param.AsString()

                        if valor and len(str(valor)) < 100:
                            params[nombre] = str(valor)
                    except:
                        pass

            if params:
                info[u"ParÃ¡metros"] = dict(list(params.items())[:10])

            elementos.append(info)

        return elementos

    def on_consultar(self, sender, args):
        """Maneja el click del botÃ³n"""
        if self.contador_consultas >= MAX_CONSULTAS_DIA:
            self.respuesta_box.Text = (
                u"âš ï¸ LÃMITE DIARIO ALCANZADO ({} consultas)\n\n"
                u"Por seguridad y control de costos, se ha alcanzado el lÃ­mite.\n"
                u"Intenta maÃ±ana o contacta al administrador."
            ).format(MAX_CONSULTAS_DIA)
            return

        if API_KEY == "TU_API_KEY_AQUI" or not API_KEY or len(API_KEY) < 20:
            self.respuesta_box.Text = (
                u"âŒ ERROR: API Key no configurada\n\n"
                u"Edita el archivo 'api_config.json' en la carpeta del script\n"
                u"y reemplaza 'TU_API_KEY_AQUI' con tu API key real.\n\n"
                u"ObtÃ©n una API key en: https://console.anthropic.com"
            )
            return

        pregunta = self.pregunta_box.Text.strip()
        if not pregunta:
            self.respuesta_box.Text = u"âš ï¸ Por favor escribe una pregunta"
            return

        elementos = self.obtener_elementos_seleccionados()

        if not elementos:
            self.respuesta_box.Text = (
                u"âš ï¸ No hay elementos seleccionados\n\n"
                u"Selecciona elementos en Revit antes de consultar."
            )
            return

        self.info_elementos.Text = u"ðŸ“¦ Analizando {} elemento(s)...".format(len(elementos))

        hash_consulta = generar_hash_consulta(pregunta, elementos)

        if hash_consulta in self.cache:
            self.respuesta_box.Text = u"ðŸ’¾ [Desde cachÃ©]\n\n" + self.cache[hash_consulta]['respuesta']
            self.info_elementos.Text = u"ðŸ“¦ {} elemento(s) | âš¡ Respuesta instantÃ¡nea (cachÃ©)".format(len(elementos))
            return

        self.respuesta_box.Text = u"â³ Consultando a Claude, espera un momento..."
        self.Dispatcher.Invoke(lambda: None)

        try:
            respuesta = self.consultar_claude(pregunta, elementos)
            self.respuesta_box.Text = respuesta
            self.contador_consultas += 1

            self.cache[hash_consulta] = {
                'respuesta': respuesta,
                'fecha': datetime.now().strftime('%Y-%m-%d')
            }
            guardar_cache(self.cache)
            guardar_contador(self.contador_consultas)
            self.actualizar_footer()

            output = script.get_output()
            output.print_md("## Consulta #{} realizada".format(self.contador_consultas))
            output.print_md("**Pregunta:** {}".format(pregunta))
            output.print_md("**Elementos:** {}".format(len(elementos)))

        except Exception as e:
            self.respuesta_box.Text = (
                u"âŒ ERROR al consultar Claude:\n\n{}\n\n"
                u"Verifica:\n"
                u"â€¢ API Key correcta\n"
                u"â€¢ ConexiÃ³n a internet\n"
                u"â€¢ CrÃ©ditos disponibles en console.anthropic.com"
            ).format(str(e))

    def consultar_claude(self, pregunta, elementos):
        """EnvÃ­a consulta a Claude API"""
        client = anthropic.Anthropic(api_key=API_KEY)

        contexto = u"InformaciÃ³n de {} elementos de Revit:\n\n".format(len(elementos))
        for i, elem in enumerate(elementos, 1):
            contexto += u"{}. {} (ID: {})\n".format(i, elem[u'CategorÃ­a'], elem['ID'])
            contexto += u"   Tipo: {}\n".format(elem['Tipo'])
            if u'ParÃ¡metros' in elem and elem[u'ParÃ¡metros']:
                contexto += u"   ParÃ¡metros clave:\n"
                for k, v in list(elem[u'ParÃ¡metros'].items())[:5]:
                    contexto += u"   - {}: {}\n".format(k, v)
            contexto += u"\n"

        mensaje = u"{}\nPregunta: {}\n\nResponde de forma concisa y prÃ¡ctica.".format(contexto, pregunta)

        modelo = "claude-haiku-4-20250408" if USAR_HAIKU else "claude-sonnet-4-5-20250929"

        message = client.messages.create(
            model=modelo,
            max_tokens=1024,
            messages=[{"role": "user", "content": mensaje}]
        )

        return message.content[0].text


def main():
    """FunciÃ³n principal"""
    ventana = ClaudeWindow()
    ventana.ShowDialog()


if __name__ == "__main__":
    main()

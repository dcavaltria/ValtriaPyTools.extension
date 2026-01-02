# -*- coding: utf-8 -*-
"""
Boton de pyRevit (IronPython 2.7) para conversar con Ephyra y ejecutar acciones
permitidas sobre la seleccion actual (mediciones, parametros, etc.).
"""

__title__ = "🜂\nEphyra"
__author__ = "Valtria"
__context__ = 'zero-doc'

import os
import json
import traceback
import sys

import clr
clr.AddReference('System.Net.Http')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import script, forms

LIB_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import valtria_lib as vlib

from Autodesk.Revit.DB import (
    Transaction,
    BuiltInParameter,
    BuiltInParameterGroup,
    StorageType
)

from System import Text
from System.Windows.Forms import (
    Form,
    Label,
    TextBox,
    Button,
    DialogResult,
    FormStartPosition,
    ScrollBars,
    AnchorStyles
)
from System.Drawing import Size, Point

from System.Net.Http import (
    HttpClient,
    HttpClientHandler,
    HttpRequestMessage,
    HttpMethod,
    StringContent
)
from System.Net.Http.Headers import MediaTypeHeaderValue

# ---------------------------------------------------------------------------
# Configuracion Anthropic
# ---------------------------------------------------------------------------
ANTHROPIC_VERSION = "2023-06-01"
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = doc.Application

# ---------------------------------------------------------------------------
# Registro basico
# ---------------------------------------------------------------------------
def info(msg):
    print(u"[INFO] {0}".format(msg))


def warn(msg):
    print(u"[WARN] {0}".format(msg))


def err(msg):
    print(u"[ERROR] {0}".format(msg))


def start_tx(name):
    tx = Transaction(doc, name)
    tx.Start()
    return tx


# ---------------------------------------------------------------------------
# Seleccion y elementos
# ---------------------------------------------------------------------------
def get_selected_elements():
    return vlib.get_selected_elements()


def get_element_type(elem):
    return vlib.get_element_type(elem)


def get_element_category_bic(elem):
    return vlib.get_element_category_bic(elem)


# ---------------------------------------------------------------------------
# Parametros
# ---------------------------------------------------------------------------
def get_param(elem, name):
    return vlib.get_param_value(elem, name)


def set_param(elem, name, value):
    return vlib.set_param_value(elem, name, value)


# ---------------------------------------------------------------------------
# Medidas basicas
# ---------------------------------------------------------------------------
def read_length(elem):
    return vlib.read_length(elem)


def read_area(elem):
    return vlib.read_area(elem)


def read_volume(elem):
    return vlib.read_volume(elem)


def measure_selection(elems):
    return vlib.measure_elements(elems)


# ---------------------------------------------------------------------------
# Parametros compartidos
# ---------------------------------------------------------------------------
def ensure_shared_parameter(param_name, param_type, group_name, categories,
                            is_instance=True, param_group=BuiltInParameterGroup.PG_DATA):
    return vlib.ensure_shared_parameter(
        param_name=param_name,
        param_type=param_type,
        group_name=group_name,
        categories=categories,
        is_instance=is_instance,
        param_group=param_group
    )


# ---------------------------------------------------------------------------
# Acciones permitidas
# ---------------------------------------------------------------------------
def action_measure_selection():
    elems = get_selected_elements()
    if not elems:
        raise Exception(u"No hay seleccion.")
    return measure_selection(elems)


def action_set_instance_param_on_selection(param_name, value):
    elems = get_selected_elements()
    if not elems:
        raise Exception(u"No hay seleccion.")
    tx = start_tx(u"Actualizar '{0}'".format(param_name))
    ok, fail = 0, 0
    try:
        for e in elems:
            try:
                set_param(e, param_name, value)
                ok += 1
            except Exception:
                fail += 1
        tx.Commit()
    except Exception:
        tx.RollBack()
        raise
    return {"updated": ok, "failed": fail}


def action_create_shared_param_for_selection(param_name="Medible", param_type="YesNo", is_instance=True):
    elems = get_selected_elements()
    if not elems:
        raise Exception(u"No hay seleccion.")
    bic = get_element_category_bic(elems[0])
    if not bic:
        raise Exception(u"No se pudo determinar la categoria.")
    ensure_shared_parameter(
        param_name=param_name,
        param_type=param_type,
        group_name="pyRevitParams",
        categories=[bic],
        is_instance=is_instance
    )
    return {"parameter": param_name, "category": str(bic), "scope": "Instance" if is_instance else "Type"}


WHITELIST_ACTIONS = set([
    "measure_selection",
    "set_instance_param",
    "ensure_shared_param",
    "export_selection",
])


def dispatch_intent(intent, dataset=None):
    action = intent.get("action")
    if action not in WHITELIST_ACTIONS:
        raise Exception(u"Accion no permitida: {0}".format(action))

    if action == "measure_selection":
        return action_measure_selection()

    if action == "set_instance_param":
        name = intent["name"]
        value = intent.get("value", "")
        return action_set_instance_param_on_selection(name, value)

    if action == "ensure_shared_param":
        name = intent.get("name", "Medible")
        ptype = intent.get("type", "YesNo")
        scope = intent.get("scope", "Instance")
        return action_create_shared_param_for_selection(name, ptype, is_instance=(scope == "Instance"))

    if action == "export_selection":
        if not dataset:
            raise Exception(u"No hay datos para exportar.")
        fmt = (intent.get("format") or "csv").lower()
        path = intent.get("path")
        export_path = vlib.export_rows(dataset, fmt, path)
        return {"export_path": export_path, "format": fmt}

    raise Exception(u"Accion desconocida: {0}".format(action))


SYSTEM_MESSAGE = (
    "Eres un asistente Revit que puede sugerir pasos y opcionalmente pedir acciones en formato JSON. "
    "Cuando necesites que ejecute algo en Revit, responde solo con un objeto JSON valido siguiendo el esquema indicado. "
    "En el resto de los casos responde en espanol de forma breve y clara."
)

ACTION_SCHEMA = (
    "Acciones permitidas (usa exactamente estas claves si devuelves JSON):\n"
    "- Medir seleccion: {\"action\":\"measure_selection\"}\n"
    "- Asegurar parametro compartido: "
    "{\"action\":\"ensure_shared_param\",\"name\":\"Texto\",\"type\":\"Text|Number|YesNo\",\"scope\":\"Instance|Type\"}\n"
    "- Fijar parametro de instancia: "
    "{\"action\":\"set_instance_param\",\"name\":\"Texto\",\"value\":\"Texto o numero\"}\n"
    "- Exportar datos de la seleccion: "
    "{\"action\":\"export_selection\",\"format\":\"csv|json|excel\",\"path\":\"ruta opcional\"}\n"
    "Si no necesitas ejecutar acciones, responde normalmente. "
    "Si ejecutas una accion, devuelve solo el JSON."
)


# ---------------------------------------------------------------------------
# Dialogo multilinea (WinForms)
# ---------------------------------------------------------------------------
class _MultiLineForm(Form):
    def __init__(self, message, title, default_text):
        Form.__init__(self)
        self.Text = title or "Entrada"
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ClientSize = Size(600, 360)

        self.label = Label()
        self.label.Text = message or ""
        self.label.AutoSize = False
        self.label.Location = Point(12, 12)
        self.label.Size = Size(576, 40)
        self.label.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.label)

        self.textbox = TextBox()
        self.textbox.Multiline = True
        self.textbox.AcceptsReturn = True
        self.textbox.AcceptsTab = True
        self.textbox.ScrollBars = ScrollBars.Vertical
        self.textbox.Location = Point(12, 60)
        self.textbox.Size = Size(576, 230)
        self.textbox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.textbox.Text = default_text or ""
        self.Controls.Add(self.textbox)

        self.ok_button = Button()
        self.ok_button.Text = "Aceptar"
        self.ok_button.DialogResult = DialogResult.OK
        self.ok_button.Location = Point(408, 308)
        self.ok_button.Size = Size(90, 30)
        self.ok_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.Controls.Add(self.ok_button)

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancelar"
        self.cancel_button.DialogResult = DialogResult.Cancel
        self.cancel_button.Location = Point(504, 308)
        self.cancel_button.Size = Size(90, 30)
        self.cancel_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.Controls.Add(self.cancel_button)

        self.AcceptButton = self.ok_button
        self.CancelButton = self.cancel_button

    def get_text(self):
        result = self.ShowDialog()
        if result == DialogResult.OK:
            return self.textbox.Text
        return None


def show_multiline_input(message, title, default_text=""):
    return _MultiLineForm(message, title, default_text).get_text()


# ---------------------------------------------------------------------------
# Cliente HTTP
# ---------------------------------------------------------------------------
def load_api_key():
    key = os.environ.get("ANTHROPIC_API_KEY")
    if key:
        return key

    cfg = os.path.join(os.environ.get("APPDATA", ""), "pyrevit_claude.json")
    if os.path.exists(cfg):
        with open(cfg, "r") as fh:
            data = json.load(fh)
        key = data.get("ANTHROPIC_API_KEY")
        if key:
            return key

    raise EnvironmentError(
        "No encuentro la API key. Define ANTHROPIC_API_KEY o %APPDATA%\\pyrevit_claude.json."
    )


try:
    API_KEY = load_api_key()
except EnvironmentError as load_err:
    forms.alert(unicode(load_err), exitscript=True)


def build_http_client():
    handler = HttpClientHandler()
    handler.UseProxy = True
    return HttpClient(handler)


HTTP_CLIENT = build_http_client()


def anthropic_request(method, url, body_dict=None):
    request = HttpRequestMessage(HttpMethod(method), url)
    request.Headers.Add("x-api-key", API_KEY)
    request.Headers.Add("anthropic-version", ANTHROPIC_VERSION)

    if body_dict is not None:
        payload = json.dumps(body_dict, ensure_ascii=False)
        content = StringContent(payload, Text.Encoding.UTF8, "application/json")
        content.Headers.ContentType = MediaTypeHeaderValue("application/json")
        request.Content = content

    response = HTTP_CLIENT.SendAsync(request).Result
    raw = response.Content.ReadAsStringAsync().Result

    if not response.IsSuccessStatusCode:
        raise Exception("HTTP {0} {1} | {2}".format(
            int(response.StatusCode), response.ReasonPhrase, raw
        ))

    if not raw:
        return None

    try:
        return json.loads(raw)
    except Exception:
        return raw


def list_models():
    data = anthropic_request("Get", "https://api.anthropic.com/v1/models")
    names = []
    for item in (data.get("data") or []):
        mid = item.get("id")
        if mid:
            names.append(mid)
    return names


def choose_best_model(available):
    preferences = [
        "claude-3-5-sonnet",
        "claude-3-5-haiku",
        "claude-3-opus",
        "claude-3-sonnet",
        "claude-3-haiku",
        "claude-3"
    ]
    for pref in preferences:
        for name in available:
            if name.startswith(pref):
                return name
    return available[0] if available else None


def claude_messages(messages, model=None, max_tokens=512, system_message=None):
    if not model:
        models = list_models()
        if not models:
            raise Exception("No hay modelos disponibles para esta API key.")
        model = choose_best_model(models)
        if not model:
            raise Exception("No se pudo determinar un modelo valido con la API key actual.")

    body = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": messages
    }
    if system_message:
        body["system"] = system_message
    data = anthropic_request("Post", "https://api.anthropic.com/v1/messages", body)

    pieces = []
    for block in (data.get("content") or []):
        if block.get("type") == "text":
            pieces.append(block.get("text", ""))
    return model, "".join(pieces)


def log_error(details):
    try:
        log_path = os.path.join(os.environ.get("APPDATA", ""), "pyrevit_claude_error.log")
        with open(log_path, "wb") as fh:
            fh.write(details.encode("utf-8", "ignore"))
    except Exception:
        pass


def extract_intent_from_response(text):
    if not text:
        return None
    cleaned = text.strip()
    if cleaned.startswith("```"):
        parts = cleaned.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("{") and part.endswith("}"):
                cleaned = part
                break
    if not cleaned.startswith("{"):
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return None
        cleaned = cleaned[start:end + 1]
    try:
        data = json.loads(cleaned)
    except Exception:
        return None
    if isinstance(data, list) and data:
        data = data[0]
    if isinstance(data, dict) and "action" in data:
        return data
    return None


# ---------------------------------------------------------------------------
# Contexto para Ephyra
# ---------------------------------------------------------------------------
def describe_element(element):
    lines = []
    category = "Sin categoria"
    try:
        if element.Category:
            category = element.Category.Name
    except Exception:
        pass
    name = ""
    try:
        name = element.Name or ""
    except Exception:
        name = ""

    lines.append("Categoria: {0}".format(category or ""))
    lines.append("Nombre: {0}".format(name or ""))
    etype = get_element_type(element)
    if etype:
        lines.append("Tipo: {0}".format(getattr(etype, "Name", "") or ""))

    captured = 0
    for param in element.Parameters:
        if captured >= 8:
            break
        try:
            d = param.Definition
            if not d or not d.Name:
                continue
            if not param.HasValue:
                continue
            stype = param.StorageType
            if stype == StorageType.String:
                value = param.AsString()
            elif stype == StorageType.Double:
                value = param.AsDouble()
            elif stype == StorageType.Integer:
                value = param.AsInteger()
            else:
                value = param.AsValueString()
            if value is None:
                continue
            text = unicode(value)
            if len(text) > 80:
                continue
            lines.append("{0}: {1}".format(d.Name, text))
            captured += 1
        except Exception:
            continue
    return "\n    ".join(lines)


def build_context(elements):
    return vlib.build_context_for_elements(elements)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    elements = get_selected_elements()
    if not elements:
        use_view = forms.alert(
            "No hay elementos seleccionados.\n\n¿Quieres usar los elementos visibles en la vista activa?",
            yes=True,
            no=True
        )
        if use_view:
            elements = vlib.get_elements_in_active_view()
        if not elements:
            forms.alert("No se encontraron elementos para procesar.")
            return

    default_question = "Describe problemas, resumen y medidas para los elementos seleccionados."
    question = show_multiline_input(
        message="Escribe tu consulta para Ephyra:",
        title="Consulta a Ephyra",
        default_text=default_question
    )
    if not question:
        return

    context_text, summary, dataset_rows = build_context(elements)
    output = script.get_output()
    output.print_md("### Contexto enviado a Ephyra")
    output.print_md("```\n{0}\n```".format(context_text))

    initial_user = (
        ACTION_SCHEMA + "\n\nContexto de la seleccion:\n" +
        context_text + "\n\nPregunta del usuario:\n" + question +
        "\n\nResponde en espanol. Si necesitas que ejecute algo, entrega solo el JSON permitido."
    )

    conversation = [
        {"role": "user", "content": initial_user}
    ]

    output.print_md("### Conversacion con Ephyra")

    try:
        turn = 1
        while True:
            output.print_md("#### Turno {0} - enviando consulta".format(turn))
            model_used, respuesta = claude_messages(
                conversation, model=None, system_message=SYSTEM_MESSAGE
            )
            conversation.append({"role": "assistant", "content": respuesta})

            output.print_md("### Modelo seleccionado")
            output.print_md(model_used)
            output.print_md("### Respuesta de Ephyra")
            output.print_md(respuesta or "(respuesta vacia)")

            intent = extract_intent_from_response(respuesta)
            if intent:
                try:
                    result = dispatch_intent(intent, dataset_rows)
                    result_json = json.dumps(result, ensure_ascii=True, indent=2)
                    output.print_md("### Resultado de la accion")
                    output.print_md("```\n{0}\n```".format(result_json))
                    conversation.append({
                        "role": "user",
                        "content": "Resultado de la accion ejecutada:\n{0}".format(result_json)
                    })
                    if intent.get("action") in ("set_instance_param", "ensure_shared_param"):
                        _, _, dataset_rows = build_context(elements)
                except Exception as action_error:
                    error_text = "No se pudo ejecutar la accion solicitada: {0}".format(action_error)
                    output.print_md("### Error ejecutando accion")
                    output.print_md(error_text)
                    conversation.append({"role": "user", "content": error_text})

            follow_up = show_multiline_input(
                message=(
                    "Modelo: {0}\n\nRespuesta de Ephyra:\n\n{1}\n\n"
                    "Escribe un seguimiento para responder (deja vacio para terminar)."
                ).format(model_used, respuesta or "(respuesta vacia)"),
                title="Responder a Ephyra",
                default_text=""
            )

            if not follow_up or not follow_up.strip():
                break

            conversation.append({"role": "user", "content": follow_up.strip()})
            turn += 1

        forms.alert("Conversacion finalizada. Revisa la salida de pyRevit.")
    except Exception as exc:
        tb_text = traceback.format_exc()
        error_text = "Error: {0}\n{1}".format(repr(exc), tb_text)
        output.print_md("### Error en la llamada a Ephyra")
        output.print_md("```\n{0}\n```".format(error_text))
        log_error(error_text)
        forms.alert("Ocurrio un error al consultar Ephyra. Revisa la consola de pyRevit.")


if __name__ == "__main__":
    main()


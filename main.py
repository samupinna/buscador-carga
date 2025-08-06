# main.py
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.clock import mainthread
from kivy.core.window import Window
import threading
import speech_recognition as sr
import pandas as pd
import os
import sys

# Tamaño para pruebas en PC
Window.size = (400, 600)

class SearchWidget(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 10
        self.spacing = 10

        # Ruta del Excel
        self.data_file = 'data.xlsx'
        if hasattr(sys, '_MEIPASS'):
            self.data_file = os.path.join(sys._MEIPASS, 'data.xlsx')

        try:
            self.conditions = self.load_all_sheets()
        except Exception as e:
            print(f"❌ Error al cargar Excel: {e}")
            self.conditions = []
            self.add_widget(Label(text="Error al cargar Excel.", color=(1,0,0,1)))

        # Campo de búsqueda
        self.search_input = TextInput(hint_text='Buscar cliente o destino...', size_hint_y=None, height=50)
        self.add_widget(self.search_input)

        # Botones
        btn_layout = BoxLayout(size_hint_y=None, height=50)
        self.voice_btn = Button(text='🎙️ Hablar')
        self.voice_btn.bind(on_press=self.start_listening)
        self.search_btn = Button(text='Buscar')
        self.search_btn.bind(on_press=self.perform_search)
        btn_layout.add_widget(self.voice_btn)
        btn_layout.add_widget(self.search_btn)
        self.add_widget(btn_layout)

        # Etiqueta de resultados
        self.results_label = Label(text='Escribe o habla para buscar...', size_hint_y=None, height=40, color=(0,0,0,1))
        self.add_widget(self.results_label)

        # Scroll para resultados
        self.scroll_view = ScrollView()
        self.results_layout = GridLayout(cols=1, size_hint_y=None, spacing=10)
        self.results_layout.bind(minimum_height=self.results_layout.setter('height'))
        self.scroll_view.add_widget(self.results_layout)
        self.add_widget(self.scroll_view)

    def load_all_sheets(self):
        conditions = []
        try:
            excel_file = pd.ExcelFile(self.data_file)
        except Exception as e:
            print(f"❌ No se pudo abrir Excel: {e}")
            return []

        sheets = ['CABECERA SUR', 'CABECERA NORTE', 'CONTENEDORES', 'PICKING', 'ADMIN']
        for sheet_name in sheets:
            if sheet_name not in excel_file.sheet_names:
                continue
            df = pd.read_excel(self.data_file, sheet_name=sheet_name, header=None)
            df = df.fillna("").astype(str)
            cliente = "Desconocido"
            for _, row in df.iterrows():
                cells = [str(cell).strip() for cell in row]
                # Saltar filas vacías o cabeceras
                if not any(cell.strip() for cell in cells):
                    continue
                if len(cells) > 0 and cells[0] and cells[0].upper() in ["CLIENTES", "CLIENTE", "CONDICIONES"]:
                    continue
                if len(cells) > 0 and cells[0] and cells[0] not in ["", "nan"]:
                    cliente = cells[0]
                destino = cells[1] if len(cells) > 1 and cells[1] not in ["", "nan"] else "General"
                # Unir observaciones útiles
                obs = " | ".join([c for c in cells[2:] if c.strip()]) if len(cells) > 2 else ""
                if obs or destino != "General":
                    conditions.append({
                        "cliente": str(cliente),
                        "destino": str(destino),
                        "observaciones": str(obs),
                        "sheet": str(sheet_name)
                    })
        return conditions

    def start_listening(self, instance):
        self.results_layout.clear_widgets()
        self.results_label.text = '🎤 Escuchando...'
        threading.Thread(target=self.listen_audio, daemon=True).start()

    def listen_audio(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=1)
            try:
                audio = r.listen(source, timeout=5)
                text = r.recognize_google(audio, language="es-ES")
                self.update_text_input(text)
            except sr.UnknownValueError:
                self.update_text_input("No se entendió el audio.")
            except sr.RequestError:
                self.update_text_input("Error con Google Speech.")
            except Exception as e:
                self.update_text_input(f"Error: {str(e)}")

    @mainthread
    def update_text_input(self, text):
        self.search_input.text = text
        self.perform_search(None)

    def perform_search(self, instance):
        query = self.search_input.text.lower()
        if not query:
            self.show_results([])
            return
        results = [
            c for c in self.conditions
            if query in c["cliente"].lower() or query in c["destino"].lower() or query in c["observaciones"].lower()
        ]
        self.show_results(results)

    @mainthread
    def show_results(self, results):
        self.results_layout.clear_widgets()
        if not results:
            self.results_label.text = 'No se encontraron condiciones.'
            return
        self.results_label.text = f'✅ {len(results)} encontrado(s):'
        for res in results:
            text = f"[b]{res['cliente']}[/b]"
            if res['destino'] != 'General':
                text += f" → {res['destino']}"
            text += f"\n🔹 [b]Hoja:[/b] {res['sheet']}\n"
            if res['observaciones']:
                text += f"📝 {res['observaciones']}\n"
            label = Label(
                text=text,
                size_hint_y=None,
                height=120,
                text_size=(Window.width * 0.9, None),
                halign='left',
                valign='top',
                markup=True
            )
            label.bind(size=label.setter('text_size'))
            self.results_layout.add_widget(label)

class BuscadorApp(App):
    def build(self):
        return SearchWidget()

if __name__ == '__main__':
    BuscadorApp().run()
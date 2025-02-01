from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.scrollview import ScrollView
from kivy.uix.checkbox import CheckBox
from kivy.core.window import Window
from kivy.metrics import dp
import pandas as pd
import os
from kivy.utils import platform
from pathlib import Path

class ScrollableFileChooser(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        
        self.scroll_view = ScrollView(
            do_scroll_x=True,
            do_scroll_y=True,
            bar_width=dp(15),
            bar_color=[0.4, 0.4, 0.4, 0.9],
            bar_inactive_color=[0.5, 0.5, 0.5, 0.5],
            bar_margin=dp(2),
            scroll_type=['bars', 'content'],
            smooth_scroll_end=10
        )

        self.file_container = BoxLayout(
            orientation='vertical',
            size_hint_y=None
        )
        self.file_container.bind(minimum_height=self.file_container.setter('height'))

        self.file_chooser = FileChooserListView(
            size_hint_y=None,
            height=dp(500),
            filters=['*.xlsx', '*.xls'],
            show_hidden=False,
            multiselect=True
        )
        
        self.file_container.add_widget(self.file_chooser)
        self.scroll_view.add_widget(self.file_container)
        self.add_widget(self.scroll_view)

class ExcelCompilerLayout(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = dp(10)
        self.spacing = dp(10)
        
        self.dir_label = Label(
            text='Sélectionnez le répertoire contenant les fichiers Excel:',
            size_hint_y=None, 
            height=dp(30)
        )
        self.add_widget(self.dir_label)
        
        self.scrollable_chooser = ScrollableFileChooser(size_hint_y=0.7)
        self.file_chooser = self.scrollable_chooser.file_chooser
        self.file_chooser.path = self._get_initial_path()
        
        if platform == 'win':
            self.file_chooser.rootpath = 'C:\\'
        
        self.add_widget(self.scrollable_chooser)
        
        self.select_files_checkbox = CheckBox()
        self.select_files_label = Label(text='Sélectionner manuellement les fichiers')
        
        checkbox_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40))
        checkbox_layout.add_widget(self.select_files_checkbox)
        checkbox_layout.add_widget(self.select_files_label)
        self.add_widget(checkbox_layout)
        
        self.up_button = Button(
            text='Remonter au répertoire parent',
            size_hint_y=None,
            height=dp(40)
        )
        self.up_button.bind(on_press=self.go_to_parent_dir)
        self.add_widget(self.up_button)
        
        self.output_label = Label(
            text='Nom du fichier de sortie:',
            size_hint_y=None, 
            height=dp(30)
        )
        self.add_widget(self.output_label)
        
        self.output_name = TextInput(
            text='fichier_compile.xlsx',
            multiline=False,
            size_hint_y=None, 
            height=dp(40)
        )
        self.add_widget(self.output_name)
        
        self.compile_button = Button(
            text='Compiler les fichiers',
            size_hint_y=None,
            height=dp(50)
        )
        self.compile_button.bind(on_press=self.compile_files)
        self.add_widget(self.compile_button)
        
        self.status_label = Label(
            text='',
            size_hint_y=None,
            height=dp(30),
            text_size=(Window.width - dp(20), None),
            halign='left'
        )
        self.add_widget(self.status_label)

    def _get_initial_path(self):
        if platform == 'win':
            return 'C:\\'
        elif platform == 'linux':
            return str(Path.home())
        elif platform == 'macosx':
            return '/Users'
        return os.path.expanduser('~')

    def go_to_parent_dir(self, instance):
        current_path = self.file_chooser.path
        parent_path = os.path.dirname(current_path)
        if os.path.exists(parent_path):
            self.file_chooser.path = parent_path

    def compile_files(self, instance):
        try:
            selected_dir = self.file_chooser.path
            if not selected_dir or not os.path.isdir(selected_dir):
                self.status_label.text = 'Erreur: Veuillez sélectionner un répertoire valide'
                return

            if self.select_files_checkbox.active:
                excel_files = self.file_chooser.selection
            else:
                excel_files = [os.path.join(selected_dir, f) for f in os.listdir(selected_dir) 
                               if f.endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                self.status_label.text = 'Erreur: Aucun fichier Excel sélectionné'
                return

            all_data = []
            for file_path in excel_files:
                df = pd.read_excel(file_path)
                all_data.append(df)

            combined_df = pd.concat(all_data, ignore_index=True)
            output_path = os.path.join(selected_dir, self.output_name.text)
            combined_df.to_excel(output_path, index=False)

            self.status_label.text = f'Compilation réussie! Fichier créé: {self.output_name.text}'
        except Exception as e:
            self.status_label.text = f'Erreur lors de la compilation: {str(e)}'

class ExcelCompilerApp(App):
    def build(self):
        Window.size = (800, 600)
        return ExcelCompilerLayout()

if __name__ == '__main__':
    ExcelCompilerApp().run()

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.scrollview import ScrollView
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
        
        # Création du ScrollView
        self.scroll_view = ScrollView(
            do_scroll_x=True,
            do_scroll_y=True,
            bar_width=dp(15),
            bar_color=[0.4, 0.4, 0.4, 0.9],  # Couleur grise pour la barre
            bar_inactive_color=[0.5, 0.5, 0.5, 0.5],  # Couleur quand inactive
            bar_margin=dp(2),  # Marge pour éviter que la barre ne touche le bord
            scroll_type=['bars', 'content'],  # Permet le défilement par clic sur le contenu
            smooth_scroll_end=10  # Rend le défilement plus fluide
        )

        # Création du conteneur pour le FileChooser
        self.file_container = BoxLayout(
            orientation='vertical',
            size_hint_y=None
        )
        self.file_container.bind(minimum_height=self.file_container.setter('height'))

        # Configuration du FileChooser
        self.file_chooser = FileChooserListView(
            size_hint_y=None,
            height=dp(500),  # Hauteur fixe pour forcer le scroll
            filters=['*.xlsx', '*.xls'],
            show_hidden=False
        )
        self.file_chooser.dirselect = True

        # Ajout du FileChooser dans le conteneur
        self.file_container.add_widget(self.file_chooser)
        
        # Ajout du conteneur dans le ScrollView
        self.scroll_view.add_widget(self.file_container)
        
        # Ajout du ScrollView dans le BoxLayout principal
        self.add_widget(self.scroll_view)

class ExcelCompilerLayout(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = dp(10)
        self.spacing = dp(10)
        
        # Zone de sélection du répertoire
        self.dir_label = Label(
            text='Sélectionnez le répertoire contenant les fichiers Excel:',
            size_hint_y=None, 
            height=dp(30)
        )
        self.add_widget(self.dir_label)
        
        # Création du composant ScrollableFileChooser
        self.scrollable_chooser = ScrollableFileChooser(size_hint_y=0.7)
        self.file_chooser = self.scrollable_chooser.file_chooser
        initial_path = self._get_initial_path()
        self.file_chooser.path = initial_path
        
        if platform == 'win':
            self.file_chooser.rootpath = 'C:\\'
        
        self.add_widget(self.scrollable_chooser)
        
        # Bouton pour remonter au répertoire parent
        self.up_button = Button(
            text='Remonter au répertoire parent',
            size_hint_y=None,
            height=dp(40)
        )
        self.up_button.bind(on_press=self.go_to_parent_dir)
        self.add_widget(self.up_button)
        
        # Zone de saisie du nom du fichier de sortie
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
        
        # Bouton de compilation
        self.compile_button = Button(
            text='Compiler les fichiers',
            size_hint_y=None,
            height=dp(50)
        )
        self.compile_button.bind(on_press=self.compile_files)
        self.add_widget(self.compile_button)
        
        # Label pour les messages de statut
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

            excel_files = [f for f in os.listdir(selected_dir) 
                         if f.endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                self.status_label.text = 'Erreur: Aucun fichier Excel trouvé'
                return

            all_data = []
            for file in excel_files:
                file_path = os.path.join(selected_dir, file)
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
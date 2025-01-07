import customtkinter as ctk
import pandas as pd
import threading
import json
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class AttendanceApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Controle de Presenças")
        self.geometry("350x200")
        self.regions = self.load_regions_from_file("regions.json")

        self.region_options = list(self.regions.keys())

        self.region_label = ctk.CTkLabel(self, text="Região")
        self.region_label.pack(pady=5)

        self.region_dropdown = ctk.CTkComboBox(self, values=self.region_options)
        self.region_dropdown.pack(pady=10)

        self.generate_report_button = ctk.CTkButton(self, text="Relatório Presenças", command=self.generate_report)
        self.generate_report_button.pack(pady=10)

        self.loading_label = ctk.CTkLabel(self, text="Gerando Relatório...")
        self.loading_label.pack(pady=10)
        self.loading_label.pack_forget()

        self.loading_bar = ctk.CTkProgressBar(self, width=300, mode="indeterminate")
        self.loading_bar.pack(pady=10)
        self.loading_bar.pack_forget()

    def load_regions_from_file(self, filename):
        try:
            with open(filename, 'r') as file:
                regions = json.load(file)
            return regions
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo de regiões: {e}")
            return {}

    def load_data(self, file):
        try:
            return pd.read_excel(file)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo: {e}")
            return None

    def generate_report(self):
        def report_thread():
            self.loading_label.pack()
            self.loading_bar.pack()
            self.loading_bar.start()

            selected_region = self.region_dropdown.get()

            if not selected_region:
                messagebox.showwarning("Aviso", "Selecione uma região!")
                self.loading_bar.stop()
                self.loading_label.pack_forget()
                self.loading_bar.pack_forget()
                return

            if selected_region not in self.regions:
                messagebox.showwarning("Aviso", "Região não encontrada!")
                self.loading_bar.stop()
                self.loading_label.pack_forget()
                self.loading_bar.pack_forget()
                return

            group_names = self.regions[selected_region]
            
            file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
            if not file:
                self.loading_bar.stop()
                self.loading_label.pack_forget()
                self.loading_bar.pack_forget()
                return

            data = self.load_data(file)
            if data is None or data.empty:
                self.loading_bar.stop()
                self.loading_label.pack_forget()
                self.loading_bar.pack_forget()
                return

            filtered_data = data[data['Grupo'].isin(group_names)]
            if filtered_data.empty:
                messagebox.showinfo("Sem Dados", f"Nenhuma presença encontrada para os grupos {', '.join(group_names)}.")
                self.loading_bar.stop()
                self.loading_label.pack_forget()
                self.loading_bar.pack_forget()
                return

            invalid_permissions = [
                "Aqui você pode colocar as permissões na qual você considera que o colaborador não está presente como 'Falta', 'Férias' etc... "
            ]

            group_counts = {}

            for group in group_names:
                group_data = filtered_data[filtered_data['Grupo'] == group]
                unique_surnames = group_data['Sobrenomes'].unique()

                valid_collaborators = sum(
                    1 for surname in unique_surnames 
                    if not group_data[group_data['Sobrenomes'] == surname]['Permissão'].isin(invalid_permissions).any()
                )

                group_counts[group] = valid_collaborators

            report_data = pd.DataFrame(group_counts.items(), columns=["LOJA", "QUANTIDADE DE COLABORADORES SEM FALTAS"])
            total_collaborators = report_data["QUANTIDADE DE COLABORADORES SEM FALTAS"].sum()

            total_row = pd.DataFrame([["Total", total_collaborators]], columns=["LOJA", "QUANTIDADE DE COLABORADORES SEM FALTAS"])
            report_data = pd.concat([report_data, total_row], ignore_index=True)

            filename = f"relatorioColaboradoresSemFalta_{selected_region}.xlsx"

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=filename)
            if save_path:
                try:
                    report_data.to_excel(save_path, index=False)
                    messagebox.showinfo("Sucesso", f"Relatório salvo em {save_path}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")
            else:
                messagebox.showwarning("Aviso", "O Relatório não foi salvo.")

            self.loading_bar.stop()
            self.loading_label.pack_forget()
            self.loading_bar.pack_forget()

        thread = threading.Thread(target=report_thread)
        thread.start()

if __name__ == "__main__":
    app = AttendanceApp()
    app.mainloop()
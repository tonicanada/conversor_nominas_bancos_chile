from tkinter import *
from tkinter import ttk, messagebox, filedialog
import conversor_nominas_bancos_chile.bank_functions as bank_functions
from pathlib import Path
import pandas as pd


root = Tk()
root.title("Conversor nóminas banco")
root.geometry("300x480")
root.resizable(0, 0)
frm = ttk.Frame(root, padding=10)


# Diccionario donde se indican los diferentes formatos de pago
# a los que se puede transformar la nómina
formatobanco_dict = {
    "Banco Chile (Pagos Masivos)": {
        "banco_codigo": 1
    },
    "Banco Chile (Transf. Masivas)": {
        "banco_codigo": 1
    },
    "Santander (Transf. Masivas)": {
        "banco_codigo": 37
    },
    "BICE (nóminas)": {
        "banco_codigo": 28
    }

}


# FUNCIONES

def btn_help_onclick():
    """
    Esta función es un botón dentro del menú Tkinter que explica al usuario lo que hace
    la interfaz.
    """
    messagebox.showinfo("Info Programa", f"""
    Este programa convierte nóminas de pago en formato BCI al formato de otros bancos.
    """)


def verificar_columnas(df, columnas):
    """
    Verifica si un DataFrame tiene exactamente las columnas especificadas.

    :param df: El DataFrame a verificar.
    :param columnas: Una lista con los nombres de las columnas que se espera que tenga el DataFrame.
    :return: True si el DataFrame tiene exactamente las columnas especificadas, False de lo contrario.
    """
    # Obtener los nombres de las columnas del DataFrame
    columnas_df = list(df.columns)

    # Ordenar las listas de columnas para facilitar la comparación
    columnas.sort()
    columnas_df.sort()

    # Verificar si las columnas son iguales
    return columnas == columnas_df


def update_combobox_values_by_function(field, function, *args):
    """
    Esta función actualiza los valores de una combolist según los
    resultados de otra función que se pasa como argumento.
    """
    try:
        field['values'] = function(*args)
    except ValueError as e:
        field.set('')
        field['values'] = []
        error_msg = e.args[0]
        messagebox.showerror(
            "error", error_msg)


def update_rut_on_razonsocial_select(event):
    """
    Esta función actualiza el campo rut del menú cada vez que el usuario
    actualiza el campo 'razon_social'. Además, también actualiza el dropdown de 
    'convenios_empresa_pagosmasivos_bancochile'.
    """
    selected = event.widget.get()
    rut = bank_functions.get_rut_from_razonsocial(
        selected, entry_path_to_datosempresas.get())

    entry_rutempresa.config(state='normal')
    entry_rutempresa.delete(0, END)
    entry_rutempresa.insert(0, rut)
    entry_rutempresa.config(state='disabled')
    combobox_conveniosempresa.set("")
    update_combobox_values_by_function(combobox_conveniosempresa, bank_functions.get_conveniosbanco_pagosmasivos_bancochile_from_rut,
                                       entry_rutempresa.get(), formatobanco_dict["Banco Chile (Pagos Masivos)"]['banco_codigo'], entry_path_to_datosempresas.get())


def check_if_razonsocial_is_selected(event):
    """
    Esta función chequea que el usuario haya seleccionado el campo `razon_social'
    antes de seleccionar el formato de pago.
    """
    if entry_rutempresa.get() == '':
        messagebox.showerror(
            "error", "Favor seleccionar Razón Social antes que el formato de pago")
        combobox_formatorequerido.set('')
        return


def add_convenios_empresa_pagosmasivos_bancochile(event):
    """
    Esta función hace aparecer el campo dropdown de 'convenios_pagos_masivos_bancochile' si es que
    el usuario ha seleccionado ese método de pago.
    """
    selected = event.widget.get()
    if selected == "Banco Chile (Pagos Masivos)":
        label_conveniosempresa.place(relx=0.5, y=350, anchor="center")
        combobox_conveniosempresa.place(
            relx=0.5, y=380, anchor="center", width=200)
        update_combobox_values_by_function(combobox_conveniosempresa, bank_functions.get_conveniosbanco_pagosmasivos_bancochile_from_rut,
                                           entry_rutempresa.get(
                                           ), formatobanco_dict["Banco Chile (Pagos Masivos)"]['banco_codigo'],
                                           entry_path_to_datosempresas.get())

    elif selected != "Banco Chile (Pagos Masivos)":
        label_conveniosempresa.place_forget()
        combobox_conveniosempresa.place_forget()


def get_razonsociallist(path):
    try:
        df = pd.read_excel(Path(path))
        if verificar_columnas(df, ['razonsocial',
                                   'razonsocial_abreviatura',
                                   'rut',
                                   'banco_codigo',
                                   'cuenta_num',
                                   'convenios_pagos_masivos_bancochile']):
            razonsocial_list = list(
                df.razonsocial.drop_duplicates().to_numpy())
            return razonsocial_list
        else:
            raise AttributeError(
                'Nombre y cantidad de columnas no coincide.')
    except AttributeError as e:
        messagebox.showerror(
            "error", f"El archivo excel con los datos de las empresas no tiene el formato correcto. {e.args[0]}")


def btn_browsefile_datosempresas(entry_label):
    """
    Esta función abre una ventana en Tkinter donde se puede seleccionar un archivo excel '.xlsx' o '.xls'.
    Una vez seleccionado actualiza el campo 'path_excel_datosempresas' del menu.
    """
    try:
        fname = filedialog.askopenfilename(
            filetypes=(("Excel file", "*.xlsx"), ("Excel file", "*.xls"), ("all files", "*.*")))
        entry_label.delete(0, END)
        entry_label.insert(0, fname)
        update_combobox_values_by_function(
            combobox_razonsocial, get_razonsociallist, entry_label.get())
    except:
        raise ValueError


def btn_browsefile_inputpath(entry_label):
    """
    Esta función abre una ventana en Tkinter donde se puede seleccionar un archivo excel '.xlsx' o '.xls'.
    Una vez seleccionado actualiza el campo 'input_path' del menu.
    """
    try:
        fname = filedialog.askopenfilename(
            filetypes=(("Excel file", "*.xlsx"), ("Excel file", "*.xls"), ("all files", "*.*")))
        columnas_bci = bank_functions.get_headers_nomina_by_bankformat(
            "BCI", bank_functions.dict_encabezados_nominas_banco)
        if (verificar_columnas(pd.read_excel(fname), columnas_bci)):
            entry_label.delete(0, END)
            entry_label.insert(0, fname)
        else:
            messagebox.showerror(
                "error", "Nombre y cantidad de columnas del excel no coincide con el requerido por el BCI.")
    except:
        pass


def check_if_company_has_bankaccount(rut_empresa, formato_requerido, path_to_datosempresas):
    """
    Esta función chequea en el archivo excel 'datos_empresas.xlsx' si hay alguna cuenta asociada al banco y rut de empresa
    seleccionados. Levanta un error si no encuentra ninguna cuenta bancaria asociada al Rut, o si encuentra más de una.
    """
    banco_codigo = formatobanco_dict[formato_requerido]["banco_codigo"]
    try:
        bank_functions.get_bankaccount_from_rut_and_bancocodigo(
            rut_empresa, banco_codigo, path_to_datosempresas)
        return True
    except:
        messagebox.showerror(
            "error", "Posible causa 1: No se ha encontrado cuenta bancaria asociada para el banco y la empresa seleccionados \n\nPosible causa 2: Hay más de una cuenta bancaria para el banco y empresa seleccionados")
        return False


def btn_execution_function(path, rut_empresa, formato_requerido):
    """
    Esta función es la que llama a la función correspondiente del archivo 'bank_functions.py'
    para obtener la nómina en la que el usuario ha indicado que necesita.
    """
    try:
        path = Path(path)
        path_to_datosempresas = Path(entry_path_to_datosempresas.get())
        razonsocial_abreviatura = bank_functions.get_razonsocial_abreviatura_from_rut(
            rut_empresa, path_to_datosempresas)
        if check_if_company_has_bankaccount(rut_empresa, formato_requerido, path_to_datosempresas):
            if formato_requerido == "Banco Chile (Pagos Masivos)":
                bank_functions.bci_to_bancochile_pagosmasivos(path,
                                                              rut_empresa,
                                                              razonsocial_abreviatura,
                                                              combobox_conveniosempresa.get()[
                                                                  :3],
                                                              combobox_conveniosempresa.get()[
                                                                  11:16]
                                                              )
            elif formato_requerido == "Banco Chile (Transf. Masivas)":
                bank_functions.bci_to_bancochile_nomina_transferencias(
                    path, rut_empresa, razonsocial_abreviatura, path_to_datosempresas)
            elif formato_requerido == "Santander (Transf. Masivas)":
                bank_functions.bci_to_santander_transferenciasmasivas(
                    path, rut_empresa, razonsocial_abreviatura, path_to_datosempresas)
            elif formato_requerido == "BICE (nóminas)":
                bank_functions.bci_to_bice_nomina(
                    path, razonsocial_abreviatura,)
            messagebox.showinfo("Info", "Planilla generada correctamente.")
    except Exception as e:
        messagebox.showerror(
            "Error", f"Favor revise que estén completos todos los campos.")


def get_bottom_coordinate_from_widget(widget):
    x, y, w, h = [int(i) for i in widget.winfo_geometry().split("+")[1:]]

    # Calculamos la esquina inferior derecha de la etiqueta
    x2, y2 = x + w, y + h


# ELEMENTOS QUE ARMAN LA INTERFAZ TKINTER
# Help button
btn_help = ttk.Button(root, text="¿Qué hace este programa?",
                      command=btn_help_onclick)
btn_help.place(relx=0.5, rely=0.1, anchor=CENTER)

# Input path
btn_inputpath = ttk.Button(root, text="Ruta Archivo Input",
                           command=lambda: btn_browsefile_inputpath(entry_inputpath))
btn_inputpath.place(relx=0.5, y=100, anchor="center")
entry_inputpath = ttk.Entry(master=root, textvariable='')
entry_inputpath.place(relx=0.5, y=120, anchor="center", width=200)


# Ruta hacia el archivo excel 'datos_empresas'
btn_path_to_datosempresas = ttk.Button(root, text="Ruta Archivo Excel Datos Empresas",
                                       command=lambda: btn_browsefile_datosempresas(entry_path_to_datosempresas))
btn_path_to_datosempresas.place(relx=0.5, y=150, anchor="center")
entry_path_to_datosempresas = ttk.Entry(master=root, textvariable='')
entry_path_to_datosempresas.place(relx=0.5, y=170, anchor="center", width=200)


# Razon social empresa
label_razonsocial = ttk.Label(root, text="Razón Social Empresa origen")
label_razonsocial.place(relx=0.5, y=200, anchor="center")

combobox_razonsocial = ttk.Combobox(
    root, state="readonly", values=[])
# combobox_razonsocial.pack()
combobox_razonsocial.place(relx=0.5, y=220, anchor="center", width=200)
combobox_razonsocial.bind('<<ComboboxSelected>>',
                          update_rut_on_razonsocial_select)


# RUT empresa
label_rutempresa = ttk.Label(root, text="Rut Empresa, formato (11111111-1)")
label_rutempresa.place(relx=0.5, y=250, anchor="center")
entry_rutempresa = ttk.Entry(master=root, textvariable='76407152-2')
entry_rutempresa.place(relx=0.5, y=270, anchor="center", width=200)
entry_rutempresa.config(state=DISABLED)

# Formato requerido
label_formatorequerido = ttk.Label(root, text="Formato requerido")
label_formatorequerido.place(relx=0.5, y=300, anchor="center")
formatos = [key for key in formatobanco_dict]
combobox_formatorequerido = ttk.Combobox(
    root, state="readonly", values=formatos)
combobox_formatorequerido.place(relx=0.5, y=320, anchor="center", width=200)
combobox_formatorequerido.bind('<<ComboboxSelected>>',
                               lambda event: (check_if_razonsocial_is_selected(event),
                                              add_convenios_empresa_pagosmasivos_bancochile(event)))

# Execution button
btn_execution = ttk.Button(root, text="EJECUTAR", command=lambda: btn_execution_function(
    entry_inputpath.get(), entry_rutempresa.get(), combobox_formatorequerido.get()))
btn_execution.place(relx=0.5, y=430, anchor="center")


# Etiqueta y lista dropdown de los convenios (aplica solamente con el formato de
# pagos masivos del Banco de Chile)
label_conveniosempresa = ttk.Label(root, text="Convenios Banco")
combobox_conveniosempresa = ttk.Combobox(
    root, state="readonly")


root.mainloop()

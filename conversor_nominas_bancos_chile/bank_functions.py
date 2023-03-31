import pandas as pd
import numpy as np
import re
import json
from datetime import date
import glob
import unicodedata
from pathlib import PurePosixPath


import os
path_abs = os.path.dirname(__file__)


# Importa como diccionario los datos de los bancos de Chile
# with open(os.path.join(path_abs, 'bancos_codigos.json'), encoding="utf-8") as f:
#     bancos_codigos = json.load(f)

with open(os.path.join(os.path.dirname(__file__), 'bancos_codigos.json')) as f:
    bancos_codigos = json.loads(f.read())

# Importa como diccionario los distintos encabezados que tienen las nóminas de los diferentes bancos de Chile.
# with open(os.path.join(path_abs, 'bancos_headers_nomina.json'), encoding="utf-8") as f:
#     dict_encabezados_nominas_banco = json.load(f)
with open(os.path.join(os.path.dirname(__file__), 'bancos_headers_nomina.json')) as f:
    dict_encabezados_nominas_banco = json.loads(f.read())


def get_codebank_from_bankname(bankname):
    """
    Función que recibe como input el nombre de un banco chileno y entrega como
    output su código SBIF ("https://www.cmfchile.cl/portal/principal/613/w3-propertyvalue-29006.html")
    """
    for code in bancos_codigos:
        if bancos_codigos[code]["name"] == bankname:
            return code


def get_headers_nomina_by_bankformat(bankformat, dict_headers_bankformat):
    """
    Función que retorna el encabezado de nómina a partir del formato de banco
    requerido y el diccionario de encabezados.
    Ejemplo: Si queremos obtener el encabezado del formato de "Pagos Masivos del
    Banco de Chile" ejecutaríamos la siguiente función:
        get_headers_nomina_by_bankformat("chile_masivos", dict_encabezados_nominas_banco)
        Obtendremos como output una lista con el encabezado requerido para ese formato de
        nómina de banco
    """
    header_dict = {}
    for key in dict_headers_bankformat:
        if bankformat in dict_headers_bankformat[key]:
            id = dict_headers_bankformat[key][bankformat]["id"]
            name = dict_headers_bankformat[key][bankformat]["name"]
            header_dict[id] = name
    header = [None] * len(header_dict.keys())
    for key in header_dict:
        header[key - 1] = header_dict[key]
    return header


def get_relation_columnbank_columncode(bankformat, dict_headers_bankformat):
    """
    Función que devuelve un diccionario relacionando el nombre de la columna del encabezado del banco
    con el nombre de la variable.
    Ejemplo, si ejecutamos: get_relation_columnbank_columncode("santander", dict_encabezados_nominas_banco):
        Obtendremos el siguiente diccionario:
        {
            'Cta_origen': 'cuenta_cargo',
            'Cta_destino': 'cuenta_destino',
            'Cod_banco': 'banco_destino_codigo'
            ...
;       }
        La key del diccionario es el nombre requerido por el formato Santander y el value es el nombre estándar
        que usamos como variable para ese campo.
    """
    column_relation_dict = {}
    for key in dict_headers_bankformat:
        if bankformat in dict_headers_bankformat[key]:
            name = dict_headers_bankformat[key][bankformat]["name"]
            column_relation_dict[name] = key
    return column_relation_dict


def get_relation_columncode_columnbank(bankformat, dict_headers_bankformat):
    """
    Función que devuelve un diccionario relacionando el nombre de la variable con el nombre de la columna
    del encabezado.
    Ejemplo, si ejecutamos: get_relation_columncode_columnbank("santande", dict_encabezados_nominas_banco):
        Obtendremos el siguiente diccionario:
        {
            'cuenta_cargo': 'Cta_origen',
            'cuenta_destino': 'Cta_destino',
            'banco_destino_codigo': 'Cod_banco',
            ...
        }
    """
    column_relation_dict = {}
    for key in dict_headers_bankformat:
        if bankformat in dict_headers_bankformat[key]:
            name = dict_headers_bankformat[key][bankformat]["name"]
            column_relation_dict[key] = name
    return column_relation_dict


def get_rut_from_razonsocial(razonsocial, path_to_datosempresas):
    """
    Función que devuelve la razón social a partir del RUT, según el archivo excel 'datos_empresas.xlsx'.
    Se usa en la GUI tkinter, para completar la drop-down list del campo 'razón social'.
    """
    df = pd.read_excel(path_to_datosempresas)
    df = df[(df.razonsocial == razonsocial)]
    if len(df) == 0:
        raise KeyError(
            f"No se ha encontrado rut asociado a la empresa {razonsocial}")
    rut = list(df.rut.to_numpy())[0]
    return rut


def get_razonsocial_abreviatura_from_rut(rut, path_to_datosempresas):
    """
    Función que devuelve la abreviatura de la razón social a partir del RUT, según el archivo excel 'datos_empresas.xlsx'.
    Se usa en la GUI tkinter, para que una vez seleccionada la razón social, se autocomplete el campo 'rut empresa`.
    """
    df = pd.read_excel(path_to_datosempresas)
    df = df[(df.rut == rut)]
    if len(df) == 0:
        raise KeyError(f"No se ha encontrado Razón Social asociada al rut {rut}")
    razonsocial_abrv = list(df.razonsocial_abreviatura.to_numpy())[0]
    return razonsocial_abrv


def get_bankaccount_from_rut_and_bancocodigo(rut, banco_codigo, path_to_datosempresas):
    """
    Función que devuelve el número de cuenta y el código del banco a partir del RUT y del diccionario 'bancos_codigos', según
    el archivo excel 'datos_empresas.xlsx'. Si no hay ninguna cuenta, o hay más de una, se devuelve un error.
    Se usa en la GUI tkinter para obtener internamente el número de cuenta asociado al RUT.
    """
    df = pd.read_excel(path_to_datosempresas)
    df = df[(df.rut == rut) & (df.banco_codigo == banco_codigo)]
    if len(df) == 0:
        error_msg = f"No se ha encontrado ninguna cuenta bancaria asociada al banco {bancos_codigos[banco_codigo]['name']} para la empresa {get_rut_from_razonsocial[rut]}"
        raise ValueError(error_msg, 'foo')
    elif len(df) > 1:
        error_msg = f"Hay más de una cuenta corriente del banco {bancos_codigos[banco_codigo]['name']} asociada a la empresa {get_rut_from_razonsocial[rut]}"
        raise ValueError(error_msg, 'foo')
    return df.cuenta_num.to_numpy()[0]


def get_conveniosbanco_pagosmasivos_bancochile_from_rut(rut, banco_codigo, path_to_datosempresas):
    """
    Función que devuelve una lista con los convenios del modo "pagos masivos" del Banco de Chile,
    sgún la empresa.
    """
    if banco_codigo != 1:
        return ["no_aplica"]
    else:
        df = pd.read_excel(path_to_datosempresas)
        df = df[(df.rut == rut) & (df.banco_codigo == banco_codigo)]
        if df.empty:
            return ['']

        convenios_banco = df.convenios_pagos_masivos_bancochile.astype(str).to_numpy()[
            0].split(",")
        convenios_banco = [elemento.strip() for elemento in convenios_banco]

        # if (len(convenios_banco) == 1 and convenios_banco[0] == 'nan'):
        #     return ['']

        # Chequeo de que los convenios estén bien escritos en el archivo excel de 'datos_empresas.xlsx'
        patron = r"^\d{3} - Pago [a-zA-Z]+"
        for convenio in convenios_banco:
            if not re.match(patron, convenio):
                error_msg = "Favor revisar en el excel 'datos_empresas.xlsx' que el nombre de los convenios siga el patrón '000 - Pago xxxx'"
                raise ValueError(error_msg)
        return convenios_banco


def strip_accents(s):
    """
    Función que elimina los acentos de una cadena de texto en Python.
    """
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn')


def get_bankformat_from_bciformat(df_bcinomina, bankformat):
    """
    Función que transforma un Pandas Dataframe de nómina con formato BCI al formato requerido.
    El output es un Pandas Dataframe con el encabezado correspondiente, las columnas requeridas por formato requerido 
    que no estaban en el formato BCI original quedarán como NaN, y deberán ser completadas posterior a ejecutar esta función.

    Ejemplo:
    Input:
        |   Nº Cuenta de Cargo |   Nº Cuenta de Destino |   Banco Destino |   Rut Beneficiario |   Dig. Verif. Beneficiario | Nombre Beneficiario                                | ... |
        |---------------------:|-----------------------:|----------------:|-------------------:|---------------------------:|:---------------------------------------------------|:----|
        |             61668095 |              921260632 |              49 |           76857892 |                          0 | Seguridad Integral Hammer Spa                      | ... |
        |             61668095 |              921260632 |              49 |           76857892 |                          0 | Seguridad Integral Hammer Spa                      | ... |
        |             61668095 |             2000840602 |               1 |           76894931 |                          1 | Servicios de Ingenieria Tecnológica Spa            | ... |
    Output:
        |   Cta_origen |   moneda_origen |   Cta_destino |   moneda_destino |   Cod_banco |   RUT benef. | ... |
        |-------------:|----------------:|--------------:|-----------------:|------------:|-------------:|:----|
        |     61668095 |             nan |     921260632 |              nan |          49 |          nan | ... |
        |     61668095 |             nan |     921260632 |              nan |          49 |          nan | ... |
        |     61668095 |             nan |    2000840602 |              nan |           1 |          nan | ... |

    En este caso de ejempo el formato original de BCI no contiene datos acerca de la variable "moneda_origen", "moneda_destino", por lo que quedan como NaN.
    La columna "RUT benef. también queda como NaN dado que el formato Santander requiere el RUT junto con el Dígito Verificador,
    y en el formato BCI estos aparecen como 2 columnas separadas. Todas las columnas que quedan como NaN en el output dataframe deben ser
    completadas posteriormente a ejecutar la función. 

    """
    df_columns = get_headers_nomina_by_bankformat(
        bankformat, dict_encabezados_nominas_banco)
    df_output = pd.DataFrame(columns=df_columns)

    rel_colbci_to_colcode = get_relation_columnbank_columncode(
        "BCI", dict_encabezados_nominas_banco)
    rel_colcode_to_colsantander = get_relation_columncode_columnbank(
        bankformat, dict_encabezados_nominas_banco)

    for column in df_bcinomina:
        colcode = rel_colbci_to_colcode[column]
        if bankformat in dict_encabezados_nominas_banco[colcode].keys():
            if (dict_encabezados_nominas_banco[colcode][bankformat]["name"]) in df_columns:
                df_output[rel_colcode_to_colsantander[colcode]
                          ] = df_bcinomina[column]
    return df_output


# df = pd.read_excel("./planillas_test/20230324_nominabci.xls")
# print(df.head(10).fillna("").to_markdown(index=False))
# get_bankformat_from_bciformat(df, "santander")

def bci_to_santander_transferenciasmasivas(path, rut_empresa, razonsocial_abreviatura, path_to_datosempresas):
    """
    Función que transforma una nómina en formato BCI al formato de "Transferencias Masivas" del Banco Santander.

    Parameters:
    -----------
    path : str
        Ruta hacia el excel con la nómina en formato BCI.
    rut_empresa : str
        Rut de la empresa origen que está realizando las transferencias.
    razonsocial_abreviatura : str
        Abreviatura de la razón social de la empresa que está realizando la transferencia.
    """
    df = pd.read_excel(path)

    df_santander = get_bankformat_from_bciformat(df, "santander")

    # A partir de aquí se completan las columnas que han quedado como NaN
    rel_colcode_to_colbci = get_relation_columncode_columnbank(
        "BCI", dict_encabezados_nominas_banco)
    df_santander["Cta_origen"] = str(
        get_bankaccount_from_rut_and_bancocodigo(rut_empresa, 37, path_to_datosempresas))

    df_santander["moneda_origen"] = "CLP"
    df_santander["moneda_destino"] = "CLP"
    df_santander["Glosa correo"] = df_santander["Glosa TEF"]
    df_santander["Glosa Cliente"] = df_santander["Glosa TEF"]
    df_santander["Glosa Cartola Beneficiario"] = df_santander["Glosa TEF"]
    df_santander["RUT benef."] = df[rel_colcode_to_colbci["rut_beneficiario_sin_dv"]].astype(str) + df[
        rel_colcode_to_colbci["rut_beneficiario_dv"]].astype(str).str.lower()
    df_santander.to_excel(path.parent.joinpath(
        f"{PurePosixPath(path).stem}_{razonsocial_abreviatura}_stdr.xlsx"), index=False)
    return df_santander


# df = bci_to_santander_transferenciasmasivas("./planillas_test/20230324_nominabci.xls", "./planillas_test/20230324_nominabci.xls", "762345312-2", "tecton")


def bci_to_bice_nomina(path, razonsocial_abreviatura):
    """
    Función que transforma una nómina en formato BCI al formato banco BICE.

    Parameters:
    -----------
    path : str
        Ruta hacia el excel con la nómina en formato BCI.
    rut_empresa : str
        Rut de la empresa origen que está realizando las transferencias.
    razonsocial_abreviatura : str
        Abreviatura de la razón social de la empresa que está realizando la transferencia.
    """
    df = pd.read_excel(path)
    df_bice = get_bankformat_from_bciformat(df, "bice")

    # A partir de aquí se completan las columnas que han quedado como NaN
    rel_colcode_to_colbci = get_relation_columncode_columnbank(
        "BCI", dict_encabezados_nominas_banco)
    df_bice["rut_beneficiario_con_dv"] = df[rel_colcode_to_colbci["rut_beneficiario_sin_dv"]].astype(str) + df[
        rel_colcode_to_colbci["rut_beneficiario_dv"]].astype(str).str.upper()
    df_bice["nombre_beneficiario"] = df_bice["nombre_beneficiario"].str.replace(
        r'\W+', '', regex=True).apply(strip_accents).str[:40]
    df_bice["cuenta_destino_tipo"] = 3
    df_bice["moneda_destino"] = 0
    df_bice["oficina_origen"] = 1
    df_bice["oficina_destino"] = 1
    df_bice["mensaje_destinatario"] = df_bice["mensaje_destinatario"].str.replace(
        r'\W+', '', regex=True).apply(strip_accents)
    df_bice.to_csv(path.parent.joinpath(
        f"{PurePosixPath(path).stem}_{razonsocial_abreviatura}_bice.csv"), header=False, sep=";", index=False)


def bci_to_bancochile_nomina_transferencias(path, rut_empresa, razonsocial_abreviatura, path_to_datosempresas):
    """
    Función que transforma una nómina en formato BCI al formato banco Transferencias Masivas del Banco de Chile.

    Parameters:
    -----------
    path : str
        Ruta hacia el excel con la nómina en formato BCI.
    rut_empresa : str
        Rut de la empresa origen que está realizando las transferencias.
    razonsocial_abreviatura : str
        Abreviatura de la razón social de la empresa que está realizando la transferencia.
    """
    re_express = re.compile("[^a-zA-Z.\d\s]")
    df = pd.read_excel(path)
    df_output = get_bankformat_from_bciformat(df, "chile_transmasivas")

    # A partir de aquí se completan los campos que han quedado como NaN

    df_output['banco_destino_rut_con_dv'] = df_output.apply(
        lambda row: bancos_codigos[str(row['banco_destino_rut_con_dv'])]['rut'], axis=1
    )

    df_output['tipo_operacion'] = np.where(df_output['banco_destino_rut_con_dv'] == bancos_codigos["1"]["rut"],
                                           "TEC",
                                           "TOB")

    df_output['rut_cliente'] = rut_empresa
    df_output['cuenta_cargo'] = get_bankaccount_from_rut_and_bancocodigo(
        rut_empresa, 1, path_to_datosempresas)

    df_output['rut_beneficiario_con_dv'] = df[dict_encabezados_nominas_banco['rut_beneficiario_sin_dv']['BCI']['name']].astype(str) + \
        df[dict_encabezados_nominas_banco['rut_beneficiario_dv']
            ['BCI']['name']].astype(str)

    df_output["tipo_abono_inmediato"] = " "
    df_output["notificacion_por_email"] = "1"
    df_output["asunto_email"] = df_output["motivo_transferencia"]
    df_output["cuenta_destino_tipo"] = np.where(
        df_output['banco_destino_rut_con_dv'] == bancos_codigos["1"]["rut"], "", "CTD")

    # Aquí se unen todas las columnas en generando un archivo de texto plano, según lo requerido
    # por el formato del banco

    df_output['consolidado'] = \
        df_output.tipo_operacion + \
        df_output.rut_cliente.str.replace("-", "").astype(str).str.rjust(10, "0") + \
        df_output.cuenta_cargo.astype(str).str.rjust(12, "0") + \
        df_output.rut_beneficiario_con_dv.astype(str).str.rjust(10, "0") + \
        df_output.nombre_beneficiario.str[:30].str.replace(re_express, "").astype(str).str.rjust(30, " ") + \
        df_output.cuenta_destino.astype(str).str.rjust(18, "0") + \
        df_output.banco_destino_rut_con_dv.astype(str).str.replace("-", "").str.rjust(10, "0") + \
        df_output.monto_transferencia.astype(str).str.rjust(11, "0") + \
        df_output.tipo_abono_inmediato.astype(str) + \
        df_output.motivo_transferencia.astype(str).str[:30].str.rjust(30, " ") + \
        df_output.notificacion_por_email.astype(str) + \
        df_output.asunto_email.astype(str).str[:30].str.rjust(30, " ") + \
        df_output.email_destinatario.astype(str).str[:50].str.rjust(50, " ") + \
        df_output.cuenta_destino_tipo.astype(str)

    df_output['consolidado'].to_csv(path.parent.joinpath(
        f"{PurePosixPath(path).stem}_{razonsocial_abreviatura}_chilemasivas.txt"), header=False, index=False)


def bci_to_bancochile_pagosmasivos(path, rut_empresa, razonsocial_abreviatura, convenio_pago, nombre_nomina):
    """
    Función que transforma una nómina en formato BCI al formato banco Pagos Masivos del Banco de Chile.

    Parameters:
    -----------
    path : str
        Ruta hacia el excel con la nómina en formato BCI.
    rut_empresa : str
        Rut de la empresa origen que está realizando las transferencias.
    razonsocial_abreviatura : str
        Abreviatura de la razón social de la empresa que está realizando la transferencia.
    convenio_pago : str
        Convenio de pago
    """
    df = pd.read_excel(path)

    # Generación del encabezado del archivo
    fecha_pago = date.today().strftime("%Y%m%d")
    rut_empresa = rut_empresa.replace(".", "").replace(
        "-", "").upper().rjust(9, "0")[:9]
    num_nomina = 1
    nombre_nomina = f"{fecha_pago}{nombre_nomina}".ljust(25, " ")[:25]
    codigo_moneda = "01"
    monto_total_nomina = format(df[dict_encabezados_nominas_banco["monto_transferencia"]["BCI"]["name"]].sum(),
                                '.2f').replace(
        ".", "").rjust(13, "0")
    encabezado = f"010{rut_empresa}{str(convenio_pago).rjust(3, '0')}{str(num_nomina).rjust(5, '0')}" \
                 f"{nombre_nomina}{codigo_moneda}{fecha_pago}{monto_total_nomina}{' ' * 3}{'N'}" \
                 f"{' ' * 322}{'010201'}"

    # Generación del cuerpo del archivo
    df_chile_masivos = get_bankformat_from_bciformat(df, "chile_masivos")

    # A partir de aquí se completan las columnas NaN y se aplica el formato requerido
    # por el banco
    df_chile_masivos["tipo_registro_beneficiario"] = '02'
    df_chile_masivos["rut_emisor_sin_dv"] = rut_empresa[:-1]
    df_chile_masivos["rut_emisor_dv"] = rut_empresa[-1]
    df_chile_masivos["convenio_numero"] = convenio_pago
    df_chile_masivos["num_nomina"] = num_nomina
    df_chile_masivos["medio_pago"] = df_chile_masivos.apply(
        lambda row: "01" if row["banco_destino_codigo"] == 1 else "07", axis=1
    )
    df_chile_masivos["tipo_direccion"] = "0"
    df_chile_masivos["numero_mensaje"] = df_chile_masivos.index.to_numpy() + 1
    df_chile_masivos["vale_vista_acumulado"] = "N"
    df_chile_masivos["tipo_registro_mensaje"] = "03"
    df_chile_masivos["tipo_aviso"] = "EMA"
    df_chile_masivos["correlativo_impresion"] = '000000'
    df_chile_masivos["vale_vista_virtual"] = "S"
    df_chile_masivos["monto_transferencia"] = df_chile_masivos["monto_transferencia"].astype(
        str) + "00"

    df_chile_masivos["beneficiario_str"] = \
        df_chile_masivos["tipo_registro_beneficiario"] + \
        df_chile_masivos["rut_emisor_sin_dv"].str.rjust(9, "0") + \
        df_chile_masivos["rut_emisor_dv"] + \
        df_chile_masivos["convenio_numero"].astype(str).str.rjust(3, "0") + \
        " " * 2 + \
        df_chile_masivos["num_nomina"].astype(str).str.rjust(5, "0") + \
        df_chile_masivos["medio_pago"] + \
        df_chile_masivos["rut_beneficiario_sin_dv"].astype(str).str.rjust(9, "0") + \
        df_chile_masivos["rut_beneficiario_dv"].astype(str).str.upper() + \
        df_chile_masivos["nombre_beneficiario"].str[:60].str.ljust(60, " ") + \
        df_chile_masivos["tipo_direccion"] + \
        " " * (35 + 15 + 15 + 7 + 2) + \
        df_chile_masivos["banco_destino_codigo"].astype(str).str.rjust(3, "0") + \
        df_chile_masivos["cuenta_destino"].astype(str).str.ljust(22, " ") + \
        "000" + \
        df_chile_masivos["monto_transferencia"].astype(str).str.rjust(13, "0") + \
        df_chile_masivos["mensaje_destinatario"].str[:119].str.ljust(119, " ") + \
        df_chile_masivos["numero_mensaje"].astype(str).str.rjust(4, "0") + \
        df_chile_masivos["vale_vista_acumulado"] + \
        " " * (3 + 10 + 1) + \
        df_chile_masivos["correlativo_impresion"] + \
        df_chile_masivos["vale_vista_virtual"] + \
        " " * (45)

    df_chile_masivos["mensaje_str"] = \
        df_chile_masivos["tipo_registro_mensaje"] + \
        df_chile_masivos["rut_emisor_sin_dv"].str.rjust(9, "0") + \
        df_chile_masivos["rut_emisor_dv"] + \
        df_chile_masivos["convenio_numero"].astype(str).str.rjust(3, "0") + \
        " " * 2 + \
        df_chile_masivos["num_nomina"].astype(str).str.rjust(5, "0") + \
        df_chile_masivos["numero_mensaje"].astype(str).str.rjust(4, "0") + \
        df_chile_masivos["tipo_aviso"] + \
        df_chile_masivos["email_destinatario"].str[:96].str.ljust(96, " ") + \
        df_chile_masivos["mensaje_destinatario"].str[:250].str.ljust(250, " ") + \
        " " * (2) + "000" + " " * 20

    with open(path.parent.joinpath(f"{PurePosixPath(path).stem}{razonsocial_abreviatura}chilemasivos.txt"), 'w') as f:
        f.write(encabezado)
        f.write('\n')
        for index, row in df_chile_masivos.iterrows():
            f.write(row["beneficiario_str"])
            f.write('\n')
            f.write(row["mensaje_str"])
            f.write('\n')
        f.close()


# Otras funciones relacionadas con bancos no usadas en el programa de tkinter

def validar_nominas_rechazadas_bci(input_path, output_path):
    """
    Esta función es útil para chequear rápidamente las nóminas que han salido rechazadas en el banco BCI.
    Recibe como input la ruta donde está la carpeta con todas las nóminas excel descargadas del banco.
    """
    filenames = glob.glob(input_path)
    df_array = []
    for file in filenames:
        df = pd.read_excel(file)
        nominaName = df.iloc[2, 3]
        row_header = df.index[df["Detalle de Nómina"] == "Rut Destinatario"].tolist()[
            0]
        new_header = df.iloc[row_header]
        df = df[row_header + 1:]
        df.columns = new_header
        df["NominaName"] = nominaName
        df["Monto a Pagar ($)"] = df["Monto a Pagar ($)"].str.replace(
            ".", "").astype(int)
        df = df.dropna(how='all', axis=1)
        df_array.append(df)
    combined = pd.concat(df_array, ignore_index=True, sort=False)
    combined.to_excel(
        f"{output_path}/merged.xlsx",
        header=True,
        index=False,
    )


def split_and_save_df(df, name, size, output_dir, format, header=True):
    """
    Split a df and save each chunk in a different csv/excel/txt file.
    Parameters:
        df : pandas df to be splitted
        name : name to give to the output file
        size : chunk size
        output_dir : directory where to write the divided df
        format: "csv", "txt", "xls",
        header: True or False
    """
    for i in range(0, df.shape[0], size):
        start = i
        end = min(i + size, df.shape[0])
        subset = df.iloc[start:end]
        output_dir_including_extension = output_dir.parent.joinpath(
            f"{PurePosixPath(output_dir).stem}_{start + 1}_{end}.{format}")
        if format == "xls":
            subset.to_excel(output_dir_including_extension,
                            index=False, header=header)
        elif format == "csv":
            subset.to_csv(output_dir_including_extension,
                          index=False, header=header)
        elif format == "txt":
            subset.to_csv(output_dir_including_extension,
                          index=False, header=header)

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 23 16:54:00 2020
@author: Alexandre de Sousa Machado
"""

import os
import sys

IGNORE_PROPS = ['icon','keypreview','linktopic','lockcontrols','moveable','defaultcursortype','picture','tabindex']
def my_ord(c):
    try:
        return ord(c)
    except:
        return c

def bytestr(s):
    new_s = ""
    for c in s:
        if new_s != "": new_s = new_s + ' '
        new_s = new_s + str(my_ord(c))
    return new_s

def byte_to_str(bs):
    s = ""
    for b in bs:
        s = s + chr(my_ord(b))
    return s

def strtoint(s,default=None):
    if default is None:
        return int(s,0)
    try:
        return int(s,0)
    except:
        return default

def delete_file(filename):
    try:
        os.remove(filename)
        return True
    except:
        return False

def delete_files(path):
    for r,_,f in os.walk(path):
        for archive in f:
            delete_file(os.path.join(r,archive))

def force_directory(direc, deletingFiles = False):
    if not os.path.exists(direc):
        path,_ = os.path.split(direc)
        if force_directory(path,False):
            try:
                os.mkdir(direc)
            except:
                return False
        
    if deletingFiles:
        try:
            delete_files(direc)
        except:
            pass
    
    return True

def open_vb6(base_filename, codification="iso-8859-1"):  # iso-8859-1 is the most common codification used in Brazil 10 or more years ago... :P  
    dir,archive = os.path.split(base_filename)
    formClassName,_ = os.path.splitext(archive)
    pascal_dir = os.path.join(dir,"pas")
    force_directory(pascal_dir)
    formClassUnit = "unit_" + formClassName
    unit_file = os.path.join(pascal_dir,formClassUnit + ".pas")
    lfm_file = os.path.join(pascal_dir,formClassUnit + ".lfm")
    f = open(base_filename, "r", encoding = codification)
    return f,formClassName,unit_file,lfm_file,formClassUnit

SIMB_LINE = ""
TOKEN_LINE = ""
TOKEN_SOURCE = ""
SIMB_TYPE = 'common'
TOKEN_TYPE = 'common'
SOURCE_VB = None
DICT_OBJ = {}
NUM_LINE = 0

def scan_line():
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    global NUM_LINE

    if TOKEN_TYPE == 'eof':
        return

    SIMB_LINE = TOKEN_LINE
    SIMB_TYPE = TOKEN_TYPE
    TOKEN_SOURCE = SOURCE_VB.readline()
    TOKEN_LINE = TOKEN_SOURCE.strip()
    NUM_LINE += 1

    if TOKEN_SOURCE == "":
        TOKEN_TYPE = 'eof'
    elif TOKEN_LINE.lower().startswith('begin '):
        TOKEN_TYPE = 'obj'
    elif TOKEN_LINE.lower().startswith('beginproperty '):
        TOKEN_TYPE = 'prop'
    elif TOKEN_LINE.lower() == 'endproperty':
        TOKEN_TYPE = 'endp'
    elif TOKEN_LINE.lower() == 'end':
        TOKEN_TYPE = 'end'
    else:
        TOKEN_TYPE = 'common'
    return

def wait_token(token_type):
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    global NUM_LINE

    if TOKEN_TYPE == token_type:
        scan_line()
    else:
        raise Exception(f'ERROR - Line {NUM_LINE} - Expected token = "{token_type}"')

def analyse_property():
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    global DICT_OBJ

    wait_token('prop')
    
    while TOKEN_TYPE != 'endp':
        if TOKEN_TYPE == 'prop':
            analyse_property()
            continue
        if TOKEN_TYPE == 'obj':
            analyse_object()
            continue
        scan_line()
    
    wait_token('endp')
    return

def obtain_property(s):
    splitted_s = s.split('=')
    if len(splitted_s) < 2:
        return [], s

    k = splitted_s[0].strip().lower()
    if '(' in k or '.' in k:
        k = k.replace(')','').replace('(','.')
        k = k.split('.')
    else:
        k = [k]
    v = '='.join(splitted_s[1:]).strip()
    state = 0                        # normal
    r = ""
    for i in range(len(v)):
        if state == 0:
            if v[i] == '\'':
                break
            elif v[i] == '"':
                r += "'"
                state = 1
            else:
                r += v[i]
        elif state == 1:
            if v[i] == '"':
                r += "'"
                state = 0
            elif v[i] == '\'':
                r += "''"
            else:
                r += v[i]
    r = r.strip()
    if len(r) > 3 and r[0] == '&' and r[1] == 'H' and r[-1] == '&':
        r = '0x' + r[2:-1]
    n = strtoint(r, "int_error")
    if n == "int_error":
        return k, r
    else:
        return k, n

def analyse_object():
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    global DICT_OBJ
    global IGNORE_PROPS
    global NUM_LINE

    def include_property(dic,p_keys,p_value):
        if len(p_keys) < 0 or ".frx':" in str(p_value):
            return False
        akey = p_keys[0]
        if akey in IGNORE_PROPS:
            return False
        if len(p_keys) == 1:
            dic[akey] = p_value
        elif akey in dic:
            if not isinstance( dic[akey], dict ):
                v = dic[akey]
                dic[akey] = {'_v': v} 
            if not include_property(dic[akey], p_keys[1:], p_value):
                return False
        else:
            dic2 = {}
            if not include_property(dic2,p_keys[1:],p_value):
                return False
            dic[akey] = dic2
        return True

    wait_token('obj')

    line_piece = SIMB_LINE.split(' ')
    if len(line_piece) < 3:
        raise Exception(f'ERROR - Line {NUM_LINE} - Incomplete object definition')
    a_class_name = line_piece[1]
    a_instance_name = ' '.join(line_piece[2:])
    if a_class_name in DICT_OBJ:
        DICT_OBJ[a_class_name] += [a_instance_name]
    else:
        DICT_OBJ[a_class_name] = [a_instance_name]
    property_dic = {}
    children = []
    while TOKEN_TYPE != 'end':
        if TOKEN_TYPE == 'common':
            prop_keys, prop_value = obtain_property(TOKEN_LINE)
            include_property(property_dic,prop_keys,prop_value)
            scan_line()
            continue
        if TOKEN_TYPE == 'prop':
            analyse_property()
            continue
        if TOKEN_TYPE == 'obj':
            child = analyse_object()
            children += [child]
            continue
        raise Exception(f'ERROR - Line {NUM_LINE} - Unexpected token = "{TOKEN_TYPE}"')
    
    wait_token('end')
    if 'index' in property_dic:
        if a_class_name == 'VB.OptionButton' or a_class_name == 'VB.Label':
            a_instance_name = f"{a_instance_name}({property_dic['index']})"  # something like LblDummy(1), LblDummy(2), ...
        else: 
            print(f"WARNING: Object {a_instance_name} has index") 
            # In my VB6 sample project, only OptionButton and Label controls had Index property
            # So, the above warning is usefull to tell me that I may need to handle some other kind of control.
 
    property_dic['name'] = a_instance_name
    property_dic['_objclass'] = a_class_name
    if len(children) > 0:
        property_dic['_children'] = children
    return property_dic

def analyse_form_code():
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    code_lines = ""

    while True:
        if TOKEN_TYPE == 'common':
            code_lines += TOKEN_SOURCE
            scan_line()
            continue
        if TOKEN_TYPE == 'obj':
            obj = analyse_object()
            break

    while TOKEN_TYPE != 'eof':
        code_lines += TOKEN_SOURCE
        scan_line()

    return obj, code_lines

def analyse_form_vb6(a_filename):
    global SOURCE_VB
    global SIMB_LINE
    global TOKEN_LINE
    global TOKEN_SOURCE
    global SIMB_TYPE
    global TOKEN_TYPE
    global DICT_OBJ
    global NUM_LINE

    SOURCE_VB,a_class_name,unit_pas,unit_lfm,id_unit = open_vb6(a_filename)
    SIMB_LINE = ""
    TOKEN_LINE = ""
    TOKEN_SOURCE = ""
    SIMB_TYPE = 'common'
    TOKEN_TYPE = 'common'
    NUM_LINE = 0
    
    scan_line()
    formobj,codified = analyse_form_code()
    SOURCE_VB.close()

    lfm,pas = create_pas_and_lfm(formobj,id_unit)
    pas += "(*\n" + codified + "\n*)"

    ftxt = open(unit_lfm,"w",encoding="utf-8")
    ftxt.write(lfm)
    ftxt.close()

    ftxt = open(unit_pas,"w",encoding="utf-8")
    ftxt.write(pas)
    ftxt.close()

    return lfm,pas

def create_pas_and_lfm(formobj,id_unit):
    
    def property_value(obj,vb_obj_name,divide_15=False,isboolean=False):
        if vb_obj_name not in obj:
            return None
        v = obj[vb_obj_name]
        if isboolean:
            if int(v) == 0:
                v = 'False'
            else:
                v = 'True'
        elif divide_15:
            v = round(int(v) / 15.0)    
            # I don't know why and I didn't search for an explanation, 
            # but there's a ratio between VB6 and Pascal control 
            # dimensions (left, top, width and height).
        
        return v

    def str_pas_prop(obj,obj_level,vb_obj_name,pas_obj_name,divide_15=False,isboolean=False):
        v = property_value(obj,vb_obj_name,divide_15,isboolean)
        if v is None:
            return ""
        if divide_15 and v < -10 and pas_obj_name == 'Left':
            v += 5000
        return f"{obj_level}{pas_obj_name} = {str(v)}\n"

    def create_menu_component(obj,obj_level):
        classe_pas = 'TMenuItem'
        classe_act = 'TAction'
        nome = obj['name'].replace(')','').replace('(','__')
        nextnivel = obj_level + '  '
        olfm = f"{obj_level}object {nome}: {classe_pas}\n"
        alfm = ""
        opas = f"    {nome}: {classe_pas};\n"
        if '_children' in obj:
            olfm += str_pas_prop(obj,nextnivel,'caption','Caption')
            olfm += str_pas_prop(obj,nextnivel,'visible','Visible',isboolean=True)
            olfm += str_pas_prop(obj,nextnivel,'enabled','Enabled',isboolean=True)
            for child in obj['_children']:
                if child['_objclass'] != 'VB.Menu':
                    continue
                maislfm,maisact,maispas = create_menu_component(child,nextnivel)
                olfm += maislfm
                alfm += maisact
                opas += maispas
        else:
            nome_action = 'act_'+nome
            alfm = f"    object {nome_action}: {classe_act}\n"
            alfm += str_pas_prop(obj,"      ",'caption','Caption')
            alfm += str_pas_prop(obj,"      ",'visible','Visible',isboolean=True)
            alfm += str_pas_prop(obj,"      ",'enabled','Enabled',isboolean=True)
            alfm += f"    end\n"
            olfm += f"{nextnivel}Action = {nome_action}\n"
            opas += f"    {nome_action}: {classe_act};\n"

        olfm += f"{obj_level}end\n"
        return olfm, alfm, opas

    def create_common_control(obj,classe_pas,obj_level):
        nome = obj['name'].replace(')','').replace('(','__')
        olfm = f"{obj_level}object {nome}: {classe_pas}\n"
        nextnivel = obj_level + '  '
        olfm += str_pas_prop(obj,nextnivel,'caption','Caption')
        olfm += str_pas_prop(obj,nextnivel,'height','Height',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'left','Left',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'top','Top',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'width','Width',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'visible','Visible',isboolean=True)
        olfm += str_pas_prop(obj,nextnivel,'enabled','Enabled',isboolean=True)

        opas = f"    {nome}: {classe_pas};\n"
        if '_children' in obj:
            for child in obj['_children']:
                maislfm,maispas = create_component(child,nextnivel)
                olfm += maislfm
                opas += maispas
        olfm += f"{obj_level}end\n"
        return olfm, opas

    def create_datasource_component(obj,classe_pas,obj_level):
        nome = obj['name'].replace(')','').replace('(','__')
        olfm = f"{obj_level}object {nome}: {classe_pas}\n"
        nextnivel = obj_level + '  '
        olfm += str_pas_prop(obj,nextnivel,'left','Left',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'top','Top',divide_15=True)

        opas = f"    {nome}: {classe_pas};\n"
        if '_children' in obj:
            for child in obj['_children']:
                maislfm,maispas = create_component(child,nextnivel)
                olfm += maislfm
                opas += maispas
        olfm += f"{obj_level}end\n"
        return olfm, opas

    def create_tabsheet_control(obj,obj_level,numpage):
        nome = obj['name'].replace(')','').replace('(','__')
        nome = '_ts_' + str(numpage) + '_' + nome
        nomescroll = '_sb_' + str(numpage) + '_' + nome
        classe_scroll = 'TScrollBox'
        classe_pas = 'TTabSheet'
        nivelscroll = obj_level + '  '
        nextnivel = nivelscroll + '  '
        idx = str(numpage)
        try:
            tabcaption = obj['tabcaption'][idx]
        except:
            tabcaption = "Tab" + str(numpage)
        opas = f"    {nome}: {classe_pas};\n    {nomescroll}: {classe_scroll};\n"
        olfm = f"{obj_level}object {nome}: {classe_pas}\n{nivelscroll}Caption = {tabcaption}\n" \
            + f"{nivelscroll}object {nomescroll}: {classe_scroll}\n" \
            + f"{nextnivel}Align = alClient\n{nextnivel}BorderStyle = bsNone\n"
        try:
            nchilds = obj['tab'][idx]['controlcount']
        except:
            nchilds = 0
        nomes_childs = []
        for i in range(nchilds):
            nomechild = obj['tab'][idx]['control'][str(i)]
            if isinstance(nomechild, dict):
                nomechild = nomechild['_v']
            nomes_childs += [nomechild.replace("'","")]
        #print(f"Tab[{idx}] - Controls:",nomes_childs)
        if '_children' in obj:
            min_vtop = 100000000
            for child in obj['_children']:
                nomechild = child['name']
                if nomechild not in nomes_childs:
                    continue
                vtop = property_value(child,'top',divide_15=False)
                if vtop < 60: continue
                if vtop is not None: min_vtop = min(vtop,min_vtop)

            min_vtop -= 60
            #print(f"Min_VTOP = {min_vtop}")

            for child in obj['_children']:
                nomechild = child['name']
                if nomechild not in nomes_childs:
                    continue
                vtop = property_value(child,'top',divide_15=False)
                if vtop is not None and vtop > 60:
                    novotop = vtop - min_vtop
                    #print(f"Child {nomechild} Top: {vtop} -> {novotop}")
                    child['top'] = novotop
                maislfm,maispas = create_component(child,nextnivel)
                olfm += maislfm
                opas += maispas
        olfm += f"{nivelscroll}end\n{obj_level}end\n"
        return olfm, opas

    def create_pagecontrol_control(obj,classe_pas,obj_level):
        nome = obj['name'].replace(')','').replace('(','__')
        olfm = f"{obj_level}object {nome}: {classe_pas}\n"
        nextnivel = obj_level + '  '
        olfm += str_pas_prop(obj,nextnivel,'height','Height',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'left','Left',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'top','Top',divide_15=True)
        olfm += str_pas_prop(obj,nextnivel,'width','Width',divide_15=True)

        opas = f"    {nome}: {classe_pas};\n"
        nsheets = len(obj['tab']) if 'tab' in obj else 0
        for i in range(nsheets):
            maislfm,maispas = create_tabsheet_control(obj,nextnivel,i)
            olfm += maislfm
            opas += maispas
        olfm += f"{obj_level}end\n"
        return olfm, opas

    def create_form(obj,obj_level):
        nome = obj['name']
        if nome.lower().startswith('frm'):
            nome = nome[3:]
        elif nome.lower().startswith('form'):
            nome = nome[4:]
        nome = 'frm' + nome
        classe = 'T' + nome

        pas = f"unit {id_unit};\n\n" + \
            "{$mode objfpc}{$H+}\n\n" + \
            "interface\n\n" + \
            "uses\n" + \
            "  Classes, SysUtils, db, dbf, SdfData, sqldb, FileUtil, DateTimePicker, Forms,\n" + \
            "  Controls, Graphics, Dialogs, ComCtrls, StdCtrls, DBGrids, ExtCtrls, ExtDlgs,\n" + \
            "  Grids, Menus, MaskEdit, Buttons, ActnList;\n\n" + \
            "type\n\n  { " + classe + " }\n\n" + \
            f"  {classe} = class(TForm)\n"

        lfm = f"{obj_level}object {nome}: {classe}\n"
        nextnivel = obj_level + '  '
        lfm += str_pas_prop(obj,nextnivel,'caption','Caption')
        lfm += str_pas_prop(obj,nextnivel,'clientheight','Height',divide_15=True)
        lfm += str_pas_prop(obj,nextnivel,'clientleft','Left',divide_15=True)
        lfm += str_pas_prop(obj,nextnivel,'clienttop','Top',divide_15=True)
        lfm += str_pas_prop(obj,nextnivel,'clientwidth','Width',divide_15=True)

        if '_children' in obj:
            menunivel = nextnivel + '  '
            partelfmmenu = ""
            partelfmaction = ""
            for child in obj['_children']:
                if child['_objclass'] == 'VB.Menu':
                    maismenu, maisaction, maispas = create_menu_component(child,menunivel)
                    partelfmmenu += maismenu
                    partelfmaction += maisaction
                    pas += maispas
            if partelfmmenu != "":
                partelfmmenu = f"{nextnivel}Menu = _mnu__{nome}\n{nextnivel}object _mnu__{nome}: TMainMenu\n{partelfmmenu}{nextnivel}end\n"
                partelfmaction = f"  object _act__{nome}: TActionList\n{partelfmaction}  end\n"
                pas += f"    _mnu__{nome}: TMainMenu;\n    _act__{nome}: TActionList;\n"
                lfm += partelfmmenu + partelfmaction

            for child in obj['_children']:
                if child['_objclass'] != 'VB.Menu':
                    maislfm, maispas = create_component(child,nextnivel)
                    lfm += maislfm
                    pas += maispas

        lfm += f"{obj_level}end\n"

        pas += "  private\n" + \
            "  public\n" + \
            "  end;\n\n" + \
            "var\n" + \
            f"  {nome}: {classe};\n\n" + \
            "implementation\n\n" + \
            "{$R *.lfm}\n\n" + \
            "end.\n"

        return lfm,pas

    def create_component(obj,obj_level):
        if '_objclass' in obj:
            classe = obj['_objclass']
        else:
            return "", ""
        if classe == 'VB.Form':
            return create_form(obj,obj_level)
        if classe == 'VB.Data':
            return create_datasource_component(obj,'TDataSource',obj_level)
        if classe == 'VB.CommandButton':
            return create_common_control(obj,'TBitBtn',obj_level)
        if classe == 'TabDlg.SSTab':
            return create_pagecontrol_control(obj,'TPageControl',obj_level)
        if classe == 'MSDBGrid.DBGrid':
            return create_common_control(obj,'TDBGrid',obj_level)
        if classe == 'VB.TextBox':
            return create_common_control(obj,'TEdit',obj_level)
        if classe == 'VB.OptionButton':
            return create_common_control(obj,'TRadioButton',obj_level)
        if classe == 'VB.ComboBox':
            return create_common_control(obj,'TComboBox',obj_level)
        if classe == 'VB.Label':
            return create_common_control(obj,'TLabel',obj_level)
        if classe == 'MSComDlg.CommonDialog':
            return create_common_control(obj,'TPrintDialog',obj_level)
        if classe == 'Crystal.CrystalReport':
            return create_common_control(obj,'TPanel',obj_level)
        if classe == 'VB.Frame':
            return create_common_control(obj,'TPanel',obj_level)
        if classe == 'VB.Line':
            return create_common_control(obj,'TShape',obj_level)
        if classe == 'MSFlexGridLib.MSFlexGrid':
            return create_common_control(obj,'TStringGrid',obj_level)
        if classe == 'VB.Shape':
            return create_common_control(obj,'TShape',obj_level)
        if classe == 'VB.Menu':
            return create_menu_component(obj,obj_level)
        if classe == 'VB.Image':
            return create_common_control(obj,'TImage',obj_level)
        if classe == 'VB.CheckBox':
            return create_common_control(obj,'TCheckBox',obj_level)
        if classe == 'MSACAL.Calendar':
            return create_common_control(obj,'TDateTimePicker',obj_level)
        if classe == 'VB.Timer':
            return create_common_control(obj,'TTimer',obj_level)
        if classe == 'MSMask.MaskEdBox':
            return create_common_control(obj,'TMaskEdit',obj_level)
        return "", ""

    return create_component(formobj, '')

def main():
    global DICT_OBJ

    for i in range(1,len(sys.argv)):
        archive_name = sys.argv[i]
        print(f'VBFile = {archive_name}')
        formobj, codified = analyse_form_vb6(archive_name)
        
        #print(str(formobj))
        #print(codified)

    print([k for k in DICT_OBJ])

if __name__ == '__main__':
    main()

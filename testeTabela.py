def garantir_parametro_aux(nome_param, cat_id_int):
    """
    Garante que o parâmetro compartilhado 'nome_param' existe
    e está vinculado à categoria. Compatível com Revit 2022-2025.
    """
    cat = None
    try:
        cat = doc.Settings.Categories.get_Item(
            System.Enum.ToObject(BuiltInCategory, cat_id_int)
        )
    except:
        pass
    if cat is None:
        return False

    # Verifica se já existe
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    while it.MoveNext():
        if it.Key.Name == nome_param:
            return True

    # Cria arquivo shared param temporário
    tmp_path = os.path.join(
        os.environ.get("TEMP", "C:\\Temp"),
        "samuel_plugin_shared_params.txt"
    )
    if not os.path.exists(tmp_path):
        with open(tmp_path, "w") as f:
            f.write("# This is a Revit shared parameter file.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("GROUP\t1\tSamuelPlugin\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\n")

    old_file = None
    try:
        old_file = app.SharedParametersFilename
        app.SharedParametersFilename = tmp_path
        spf = app.OpenSharedParameterFile()
        grp = spf.Groups.get_Item("SamuelPlugin")
        if grp is None:
            grp = spf.Groups.Create("SamuelPlugin")

        # Busca definição existente no arquivo
        ext_def = None
        for d in grp.Definitions:
            if d.Name == nome_param:
                ext_def = d
                break

        if ext_def is None:
            # Tenta criar com SpecTypeId (Revit 2022+)
            try:
                from Autodesk.Revit.DB import SpecTypeId
                opts = ExternalDefinitionCreationOptions(nome_param, SpecTypeId.Number)
                ext_def = grp.Definitions.Create(opts)
            except:
                pass

            # Fallback para Revit mais antigo
            if ext_def is None:
                try:
                    from Autodesk.Revit.DB import UnitType
                    opts = ExternalDefinitionCreationOptions(nome_param, UnitType.UT_Number)
                    ext_def = grp.Definitions.Create(opts)
                except:
                    pass

            # Último fallback — só nome
            if ext_def is None:
                try:
                    ext_def = grp.Definitions.Create(nome_param)
                except:
                    pass

        if ext_def is None:
            return False

        # Vincula à categoria
        cat_set = app.Create.NewCategorySet()
        cat_set.Insert(cat)
        binding = app.Create.NewInstanceBinding(cat_set)

        # Tenta inserir com GroupTypeId (Revit 2023+)
        inserted = False
        try:
            from Autodesk.Revit.DB import GroupTypeId
            doc.ParameterBindings.Insert(ext_def, binding, GroupTypeId.Data)
            inserted = True
        except:
            pass

        # Fallback BuiltInParameterGroup para versões antigas
        if not inserted:
            try:
                from Autodesk.Revit.DB import BuiltInParameterGroup
                doc.ParameterBindings.Insert(
                    ext_def, binding, BuiltInParameterGroup.PG_DATA
                )
                inserted = True
            except:
                pass

        # Último fallback sem grupo
        if not inserted:
            try:
                doc.ParameterBindings.Insert(ext_def, binding)
            except:
                pass

        return True

    except Exception as ex:
        return False
    finally:
        if old_file is not None:
            try:
                app.SharedParametersFilename = old_file
            except:
                pass

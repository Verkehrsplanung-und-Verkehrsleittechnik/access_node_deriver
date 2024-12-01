import logging
from pathlib import Path
import numpy as np
import pandas as pd
import networkx as nx
import win32com.client as com
from helpers.Visum23Overlay import VisumOverlay


class AccessNodeDeriver:
    def __init__(self, working_directory: Path, version_filename):
        self.working_dir = working_directory
        self.gtypes_original_att = None  # placeholder for att-file where original gtypes from Visum are stored in the beginning. Is re-read before saving the finel (not debug) file.
        """Initialisiere Visum-Objekt und Filter."""
        self.Visum = self.open_visum(self.working_dir / version_filename)  # Initialisiert das Visum-Objekt
        # self.Visum.Graphic.StopDrawing = True  # Visum wird nicht unnötig bemüht den Netzeditor zu aktualisieren
        self.Visum.Filters.InitAll()  # Initialisieren aller Filter
        self.lf = self.Visum.Filters.LinkFilter()  # Erstellt einen Link-Filter
        self.nf = self.Visum.Filters.NodeFilter()  # Erstellt einen Node-Filter
        self.edit_link_atts = self.Visum.Net.CreateEditAttributePara()  # Central object for editing of link attributes
        self.edit_link_atts.SetAttValue("NETOBJECTTYPE", "LINK")
        self.intersection_nodes_df = None  # placeholder for DataFrame containing the nodes in the network that are intersection nodes (result from determine_nodetypes())
        self.access_nodes_df = None  # placeholder for DataFrame containing the access nodes as result of clustering intersection_nodes (result from cluster_by_nodetype())
        self.nodes4visum_df = None  # placeholder for DataFrame containing the access nodes in Visum format
        self.links4visum_df = None  # placeholder for DataFrame containing the helper links connecting the access nodes with the intersection nodes in each cluster in Visum format

    def open_visum(self, version_file_path=''):
        """
        open an instance of PTV Visum with the COM Interface
        :return: Visum COM object
        """
        logging.debug('initialize visum instance')
        Visum = com.Dispatch("Visum.Visum.230")
        Visum = VisumOverlay(Visum)
        Visum.Graphic.ShowMaximized()
        if version_file_path:
            Visum.LoadVersion(version_file_path)
        return Visum

    def add_uda(self, netobj, uda_name, value_type, def_val, comment=''):
        """
        Add user-defined attribute (UDA) to specified net object if not existent and add Comment

        :param netobj: net object to add the uda to
        :param uda_name: name of the UDA
        :param value_type: See Visum COM documentation "ValueType Enumeration" (e.g. 1 = Int, 2 = Real, 5 = String, 9 = Bool, ...)
        :param def_val: dafault value for the UDA
        :param comment: description of the UDA
        """

        try:
            netobj.AddUserDefinedAttribute(uda_name, uda_name, uda_name, value_type, defval=def_val)
            netobj.SetAllAttValues(uda_name, def_val, OnlyActive=False)
            netobjatts = netobj.Attributes.GetAll
            for Attr in netobjatts:
                if Attr.Code == uda_name:
                    Attr.Comment = comment
                    Attr.ValueDefault = def_val
        except:
            logging.debug(f'UDA {uda_name} might exist already.')

    def AddMat(self, ObjId, Mat_Code, Mat_Name, Objecttype, Matrixtype, Formula):
        """
        Add a matrix.
        :param ObjId: Number of the matrix to add
        :param Mat_Code: Code of the matrix (short name)
        :param Mat_Name: Name of the matrix
        :param Objecttype: Object reference type of the new matrix (Zones = 2)
        :param Matrixtype: Type of the new matrix (Demand = 3, Skim = 4)
        :param Formula: Formula if it is a formula matrix
        :return: Pointer to Matrix object for further use
        """
        if not Mat_Name:
            Mat_Name = Mat_Code
        if not Formula:
            M = self.Visum.Net.AddMatrix(ObjId, Objecttype, Matrixtype)
        else:
            M = self.Visum.Net.AddMatrixWithFormula(ObjId, Formula, Objecttype, Matrixtype)

        M.SetAttValue("Code", Mat_Code)
        M.SetAttValue("Name", Mat_Name)

        return M

    def get_multi_netobj_atts_with_id(self, netobj, atts_dict_withdtype: dict):
        """
        Adding to Visum.Net.XX.GetMultipleAttributes() this funtion also returns the Visum ID via GetMultiAttValues().
        merge on "No" --> ambiguous for links therefore needs a unique_identifier
        :param netobj: specify the net object from which attributes should be retrieved
        :param atts_dict_withdtype: dictionary with requested attributes and datatypes. Example: {"NO": int, "XCOORD": float, "YCOORD": float}
        :return: returns a DataFrame containing the requested attributes from the specified net object
        """
        # Check type of Visum net object using pywin32 access to OLE object
        netobj_type = netobj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]

        # Handling of net object that are not links
        if netobj_type != 'ILinks':
            # read attributes as a DataFrame
            df_netobj = pd.DataFrame(netobj.GetMultipleAttributes(list(atts_dict_withdtype.keys()), OnlyActive=True), columns=atts_dict_withdtype.keys()).astype(
                atts_dict_withdtype)

            # Additionally, read Visum Index (necessary to SetMultiAttValues later)
            df_netobj_index = pd.DataFrame(netobj.GetMultiAttValues("NO", OnlyActive=False), columns=["VISUM_ID", "NO"]).astype({"VISUM_ID": int, "NO": int})
            df_netobj = pd.merge(df_netobj, df_netobj_index, on='NO')
            # Place VISUM_ID in the first column
            v_id = df_netobj.pop('VISUM_ID')
            df_netobj.insert(0, 'VISUM_ID', v_id)

        # special handling of net object that are links with UDA LINK_IDENTIFIER
        else:
            # special identifier consisting of NO-FROMNODENO-TONODENO
            logging.debug(
                'Since Visum_IDs for links are ambiguos (same ID for both directions) get_multi_netobj_atts_with_id adds a UDA for links (LINK_IDENTIFIER) as NO-FROMNODENO-TONODENO.')
            self.add_uda(netobj, 'LINK_IDENTIFIER', 5, '', 'special link identifier consisting of NO-FROMNODENO-TONODENO')
            # initial setting of the identifier UDA
            x = self.Visum.Net.CreateEditAttributePara()
            x.SetAttValue("NETOBJECTTYPE", "LINK")
            x.SetAttValue("INCLUDESUBCATEGORIES", "0")
            x.SetAttValue("ONLYACTIVE", "0")
            x.SetAttValue("RESULTATTRNAME", "LINK_IDENTIFIER")
            x.SetAttValue("FORMULA", 'NUMTOSTR([NO],0) + "-" + NUMTOSTR([FROMNODENO],0) + "-" + NUMTOSTR([TONODENO],0)')
            self.Visum.Net.EditAttribute(x)

            # add the UDA to atts_dict_withdtype
            atts_dict_withdtype['LINK_IDENTIFIER'] = str

            # read attributes as a DataFrame
            df_links = pd.DataFrame(netobj.GetMultipleAttributes(list(atts_dict_withdtype.keys()), OnlyActive=True), columns=atts_dict_withdtype.keys()).astype(
                atts_dict_withdtype)

            # Additionally, read Visum Index (necessary to SetMultiAttValues later)
            df_links_index = pd.DataFrame(netobj.GetMultiAttValues("LINK_IDENTIFIER", OnlyActive=False), columns=["VISUM_ID", "LINK_IDENTIFIER"]).astype(
                {"VISUM_ID": int, "LINK_IDENTIFIER": str})
            df_links = pd.merge(df_links, df_links_index, on='LINK_IDENTIFIER')
            # Place VISUM_ID in the first column
            v_id = df_links.pop('VISUM_ID')
            df_links.insert(0, 'VISUM_ID', v_id)
            df_netobj = df_links.copy()

        return df_netobj

    def new_object_number(self, netobject):
        """
        Generates a new object number for a specified network object type in Visum.

        This function calculates the next available object number based on the highest
        existing number for the specified `netobject` type in Visum. If the maximum number
        range is reached, it attempts to find a gap within existing numbers to avoid
        exceeding the limit.

        :param netobject: The type of network object in Visum (e.g., 'Nodes', 'Links', etc.).
        :return: An integer representing the next available object number for the specified
                 network object type.
        """

        highest_netobj_no = self.Visum.Net.AttValue(rf'MAX:{netobject}\NO')
        if highest_netobj_no is not None:
            highest_netobj_no = int(highest_netobj_no)
        else:
            highest_netobj_no = 0

        power = 1
        while power < highest_netobj_no:
            power *= 10
        result = 10 * power

        # Check for Visum limit
        if result > 1000000000:
            logging.debug(f'Neue Netzobjekte des Typs {netobject} müssen zwischen vorhandenen Nummern eingefügt werden, da der Nummernbereich ausgeschöpft ist.')
            netobj = getattr(self.Visum.Net, netobject)
            netobj_nos = pd.DataFrame(netobj.GetMultiAttValues("NO", OnlyActive=False), columns=["VISUM_ID", "NO"]).astype({"VISUM_ID": int, "NO": int})
            diff = netobj_nos['NO'].diff()
            idxmax = diff.idxmax()
            start = netobj_nos.loc[idxmax - 1, 'NO']
            end = netobj_nos.loc[idxmax, 'NO']

            # find the largest power of 10 that is less than or equal to start
            power_of_10 = int(np.log10(start))
            start_gap = 10 ** (power_of_10 + 1)
            if (end - start_gap) > 25000:
                logging.debug(f'Not objects of type {netobject} will be added starting from number: {start_gap}')
                result = start_gap
            elif (end - start) > 25000:
                logging.debug(f'Not objects of type {netobject} will be added starting from number: {start}')
                result = start
            else:
                logging.debug(
                    f'Net-File will not add {netobject}. No large enough number space available. {netobject} will be written starting at {result}. Manual adjustment of the net file is necessary to enable correct reading of the file.')

        return int(result)

    def prepare_net(self, set_gtypes_osm: bool, working_dir: Path, Visum_helper_files_path: Path):
        """
        This function prepares the net, particularly the global link types (gtypes) and added link attributes.
        The gtypes must be set manually in Visum according to the specification.
        Specification: Values 1 - 5 for drivable roads (1 = major road, 5 = residential street), 77 for ramps (if known), 99 for others.
        Optionally, the gtypes can be set by using the link type names from OpenStreetMap. These must be preent in the Name attribute of the Viusm LinkTypes.

        :param set_gtypes_osm: determine whether the gtype in Visum should be set according to link type names from osm (see function set_gtypes_for_osm_import())
        :param working_dir: directory where the original gtypes should be stored as backup
        :param Visum_helper_files_path: project directory that contains the list layout 'gtypes_original.llax' to store origignal gtypes
        """
        logging.info(f'Preparation of the network: Setting global link types.')
        # Backup der originalen Obertypen erstellen
        self.gtypes_original_att = working_dir / 'gtypes_original.att'
        gtypes_original_llax = Visum_helper_files_path / 'gtypes_original.llax'
        self.Visum.IO.SaveAttributeFile(self.gtypes_original_att, ListLayoutFile=gtypes_original_llax, Separator=59)

        # Strecken die gesperrte Gegenrichtungen abbilden, sollten einem eigenen Streckentypen zugeordnet werden z. B. Typ-Nr 0
        self.lf.AddCondition("OP_NONE", False, "TSYSSET", "EqualVal", "")
        self.Visum.Net.Links.SetAllAttValues(Attribut="TYPENO", newValue=0, OnlyActive=True)
        self.lf.Init()

        def set_gtypes_for_osm_import():
            """
            Setze die Obertypen auf Streckentypen-Ebene basierend auf den Namen der Streckentypen (OSM-Key: highway).
            """
            # Im geöffneten Visum vorhandene Streckentypen einlesen
            self.Visum_linktypes_df = self.get_multi_netobj_atts_with_id(self.Visum.Net.LinkTypes, {"NO": int, "NAME": str})

            # Festlegung einer Zuordnung für OSM-Streckentypen zu gtypes
            osm_obertypen = {
                'motorway': 1,
                'trunk': 2,
                'primary': 2,
                'secondary': 3,
                'tertiary': 4,
                'unclassified': 5,
                'residential': 5,
                'road': 5,
                'living_street': 5,
                'service': 6,
                # Hinweis der Typ 6 wurde bei OSM nur der Vollständigkeit halber für "service" roads vergeben und wird nicht weiter beachtet.
            }

            # Check for keys that don't exist in the NAME column
            missing_keys = []
            for road_type in osm_obertypen.keys():
                if not self.Visum_linktypes_df['NAME'].str.contains(road_type, case=False).any():
                    missing_keys.append(road_type)
            if missing_keys:
                logging.warning(f"The following osm link types are not present in Visum LinkType Names: {', '.join(missing_keys)}. This will likely lead to wrong results. Please check the naming of link types in Visum.")

            # Neue Spalte 'GTYPE' mit 99 als Defaultwert anlegen
            self.Visum_linktypes_df['GTYPE'] = 99

            # Über Typen-dictionary iterieren und über passende strings zuordnen
            for road_type, main_type in osm_obertypen.items():
                # Use str.contains to find rows where 'NAME' contains the road type
                mask = self.Visum_linktypes_df['NAME'].str.contains(road_type, case=False)
                # Assign the main type to the rows where the mask is True
                self.Visum_linktypes_df.loc[mask, 'GTYPE'] = main_type
            keyword_rampen = '_link'
            if keyword_rampen:
                mask_link = self.Visum_linktypes_df['NAME'].str.contains(keyword_rampen, case=False)
                self.Visum_linktypes_df.loc[mask_link, 'GTYPE'] = 77
            self.Visum.Net.LinkTypes.SetMultiAttValues("GTYPE", self.Visum_linktypes_df[["VISUM_ID", "GTYPE"]].values)

        if set_gtypes_osm:
            set_gtypes_for_osm_import()
        else:
            logging.warning('Make sure that the global type (GTYPE) of the link types is set manually according to the specification')

        # Strecken-Attribut das den ungerichteten Obertyp enthält (Info: hier werden Rampen später statt 77 auf den Wert -1 gesetzt)
        self.add_uda(self.Visum.Net.Links, uda_name='ZK_OBERTYP_UNGERICHTET', value_type=1, def_val=99)
        self.edit_link_atts.SetAttValue("ONLYACTIVE", "0")
        self.edit_link_atts.SetAttValue("RESULTATTRNAME", "ZK_OBERTYP_UNGERICHTET")
        self.edit_link_atts.SetAttValue("FORMULA", 'MIN([GTYPE], [REVERSELINK\LINKTYPE\GTYPE])')
        self.Visum.Net.EditAttribute(self.edit_link_atts)

    def spatial_selection(self, territory_no):
        """
        Make a spatial selection on the territory defined by territory_no and determine nodes in the periphey of the selection.
        :param territory_no: Attribute No of the territory in Visum that should serve as space to select
        """
        logging.info(f'Spatial selection on territory number {territory_no}.')
        self.Visum.Net.SetTerritoryActive(territory_no)

        # 1.2.1 Knoten in Randlage sollten ausgeschlossen werden
        # TODO: Prüfen... Zur Bildung von Netzabschnitten ist das nicht sinnvoll, evtl. ein eigener toggler?
        # Randlage = Enden des Netzes außerhalb des Untersuchungsraums
        # Ermittle, ob ein Knoten in Randlage ist und schreibe es in ein UDA
        self.add_uda(self.Visum.Net.Nodes, 'ZK_IST_RANDKNOTEN', value_type=9, def_val=0)
        self.nf.Init()
        self.nf.AddCondition("OP_NONE", False, "COUNT:INLINKS", "NotEqualAtt", "COUNTACTIVE:INLINKS")
        self.Visum.Net.Nodes.SetAllAttValues("ZK_IST_RANDKNOTEN", 1, OnlyActive=True)

    def preprocess_roundabouts(self, roundabout_uda):
        """
        Preprocess of roundabouts if they are labeled in the data. roundabouts should be handled exactly as ramps by the criteria in determine_nodetype()
        :param roundabout_uda: boolean label to identify links in the network that are roundabouts
        """
        logging.info('Preprocessing of roundabouts.')
        if self.Visum.Net.Links.AttrExists(roundabout_uda):
            self.edit_link_atts.SetAttValue("ONLYACTIVE", "0")
            self.edit_link_atts.SetAttValue("RESULTATTRNAME", "ZK_OBERTYP_UNGERICHTET")
            # exclude construction links (many roundabouts are in construction in osm)
            if self.Visum.Net.Links.AttrExists('OSM_CONSTRUCTION'):
                self.edit_link_atts.SetAttValue("FORMULA", f'IF([{roundabout_uda}]=1 & [OSM_CONSTRUCTION]=0, 77, [ZK_OBERTYP_UNGERICHTET])')
            else:
                self.edit_link_atts.SetAttValue("FORMULA", f'IF([{roundabout_uda}]=1, 77, [ZK_OBERTYP_UNGERICHTET])')
            self.Visum.Net.EditAttribute(self.edit_link_atts)
        else:
            logging.warning(f'korrigiere_kreisverkehre ist aktiviert, aber das angegebene UDA {roundabout_uda} existiert nicht. Kreisverkehre können nicht berücksichtigt werden.')

        # Rampen Obertyp = 77 nachträglich auf -1 in ZK_OBERTYP_UNGERICHTET
        self.edit_link_atts.SetAttValue("FORMULA", 'IF([ZK_OBERTYP_UNGERICHTET] = 77, -1, [ZK_OBERTYP_UNGERICHTET])')
        self.Visum.Net.EditAttribute(self.edit_link_atts)


    def ramp_analysis(self):
        """Write the maintype of the links connected by the ramps to the ramp links
        :return: result in Visum link attribute ZK_OBERTYP_VERSCHIEDENE (with ramps) and ZK_ÜBERGABEANALLE (start/end node of the ramp)
        """
        logging.info('Processing of ramps (incl. roundabouts).')
        # Filter auf Strecken die Rampen sind, mit vorab belegtem Attribut
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "ZK_OBERTYP_UNGERICHTET", "EqualVal", -1)

        if self.Visum.Net.Links.CountActive:

            # Korrektur Rampen mit networkx über subgraphs
            # Einlesen der Rampen in einen Dataframe
            # dabei wird am Von- und Nach-Knoten noch eingelesen welche Gtypes dort außerdem noch vorkommen
            atts_dict_withdtype = {"NO": int, "FROMNODENO": int, "TONODENO": int, r"FROMNODE\DISTINCT: INLINKS\ZK_OBERTYP_UNGERICHTET": str,
                                   r"TONODE\DISTINCT: INLINKS\ZK_OBERTYP_UNGERICHTET": str}
            df_links = self.get_multi_netobj_atts_with_id(self.Visum.Net.Links, atts_dict_withdtype)

            # Verknüpfen der Attribute des Von und Nach-Knotens für jede Strecke
            df_links['TFGTYPE'] = (df_links[r"FROMNODE\DISTINCT: INLINKS\ZK_OBERTYP_UNGERICHTET"] + ',' + df_links[r"TONODE\DISTINCT: INLINKS\ZK_OBERTYP_UNGERICHTET"]).apply(
                lambda a: [int(x) for x in a.split(',') if x != ''])
            # Herausfiltern welche davon anders sind als Rampen (-1) und vom Typ <= 5 sind
            df_links['TFGTYPE_less5'] = df_links['TFGTYPE'].apply(lambda s: [x for x in s if (x <= 5) & (x > -1)])

            # Erstelle einen DiGraphen aus dem Dataframe mit links
            G = nx.from_pandas_edgelist(df_links, 'FROMNODENO', 'TONODENO', edge_attr=['VISUM_ID', 'NO', 'TFGTYPE_less5'], create_using=nx.DiGraph())
            # Ermittle, welche Strecken miteinander verbunden sind (Start der Rampe bis Ende)
            connected_components = list(nx.weakly_connected_components(G))

            # create a list to store the results for each group
            # enthält die zusammenhängenden
            grouped_links = []

            for component in connected_components:
                # create a subgraph for the current component
                subgraph = G.subgraph(component)
                # get the Visum_IDs of all edges in the subgraph
                ids = [d['VISUM_ID'] for _, _, d in subgraph.edges(data=True)]
                # get the LinkNOs of all edges in the subgraph
                nos = [d['NO'] for _, _, d in subgraph.edges(data=True)]
                # concat gtypes of the To and From Nodes and sort them
                tfgtype_lists = [d['TFGTYPE_less5'] for _, _, d in subgraph.edges(data=True)]
                tfgtype_sorted = sorted([item for sublist in tfgtype_lists for item in sublist])

                # store the result in the list
                grouped_links.append({'visum_ids': ids, 'linknos': nos, 'ftgtypes': tfgtype_sorted})

            # convert the list to a dataframe
            grouped_links_df = pd.DataFrame(grouped_links)
            # leere Einträge entfernen (passiert hauptsächlich bei im Bau befindlichen und somit isolierten Infrastrukturen)
            grouped_links_df = grouped_links_df[grouped_links_df['ftgtypes'].map(len) > 0]
            # Voneinander verschiedene vorkommende Typen extrahieren (distinct)
            grouped_links_df['ZK_OBERTYP_VERSCHIEDENE'] = grouped_links_df['ftgtypes'].apply(lambda x: ','.join([str(y) for y in set(x)]))

            # subgraphs with just one type must be removed --> type 6 (service)?
            # "lose Enden", len(edit_link_atts)=2 bedeutet es gab insgesamt nur einen vorkommenden Typen an nur einem Knoten
            mask = grouped_links_df.ftgtypes.apply(lambda x: len(x) == 2)
            grouped_links_df.loc[mask, 'ZK_OBERTYP_VERSCHIEDENE'] = '6'

            # Erstellung eines Dataframes, der die Visum_IDs mit den ermittelten gtypes enthält
            linkids2edit = grouped_links_df.explode('visum_ids')[['visum_ids', 'ZK_OBERTYP_VERSCHIEDENE']]
            # Zurückschreiben nach Visum
            self.Visum.Net.Links.SetMultiAttValues('ZK_OBERTYP_VERSCHIEDENE', linkids2edit[['visum_ids', 'ZK_OBERTYP_VERSCHIEDENE']].to_numpy())

        # Ermitteln der Übergabeknoten ans nachgeordnete Netz aus der initialen Zuordnung (Knoten-UDA "ÜbergabeAn")
        # Filtere Knoten
        self.nf.Init()
        self.nf.AddCondition("OP_NONE", False, "ONACTIVELINK", "EqualVal", 1)
        self.nf.AddCondition("OP_AND", False, "MAX:INLINKS\ZK_OBERTYP_UNGERICHTET", "GreaterVal", -1)
        if self.Visum.Net.Nodes.CountActive:
            def remove_negative(col):
                """remove the negative values"""
                rm = sorted(set(col.split(',')))
                if '-1' in rm:
                    rm.remove('-1')
                return ','.join(rm)

            # Erzeuge und befülle UDA
            # add_uda(Visum.Net.Nodes, 'ÜbergabeAn', 1, -1, ' Übergabeknoten an Minimum Netz der Ebene 0 bis 5 (int)')
            # df_nodes_ubergabe['G_min'] = df_nodes_ubergabe['ZK_OBERTYP_VERSCHIEDENE'].apply(process_column_2nd_lowest)
            # Visum.Net.Nodes.SetMultiAttValues('ÜBERGABEAN', df_nodes_ubergabe[["VISUM_ID", "G_min"]].to_numpy())

            self.add_uda(self.Visum.Net.Nodes, 'ZK_ÜBERGABEKNOTEN', 5, '', 'Übergabeknoten (Anfang/Ende einer Rampe) an alle Netze der Ebenen 0 bis 5 (text)')
            df_nodes_ubergabe = pd.DataFrame(self.Visum.Net.Nodes.GetMultiAttValues("DISTINCT:INLINKS\ZK_OBERTYP_UNGERICHTET", OnlyActive=True),
                                             columns=["VISUM_ID", "ZK_OBERTYP_VERSCHIEDENE"])
            df_nodes_ubergabe['G_all'] = df_nodes_ubergabe['ZK_OBERTYP_VERSCHIEDENE'].apply(remove_negative)
            self.Visum.Net.Nodes.SetMultiAttValues('ZK_ÜBERGABEKNOTEN', df_nodes_ubergabe[["VISUM_ID", "G_all"]].to_numpy())

    def identify_uturns(self):
        r"""Entfernt Rampen, die U-Turns sind aus der Betrachtung
        Wenn Rampenstreckenzüge kürzer als 50m sind und die damit verbunden Strecken denselben Straßennamen tragen, handelt es sich mit angrenzender Wahrscheinlichkeit um U-Turns
        Bedingung: Strecken sind mit Straßennamen benannt (LINK\NAME)

        Output in das Visum-Strecken-Attribut "ZK_IS_UTURN"
        """
        logging.info('Identification and processing of u-turns.')
        # Führe ein Knotenattribut in Visum ein, das angibt, ob es sich um einen U-Turn handelt
        self.add_uda(self.Visum.Net.Links, 'ZK_IS_UTURN', value_type=9, def_val=False)
        self.Visum.Net.Links.SetAllAttValues('ZK_IS_UTURN', False, OnlyActive=False)

        # Filtere auf Strecken, die Rampen sind und lediglich über einen Obertypen entsprechend rampen_obertyp_analyse() haben also gleichrangige Strecken verbinden
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "LINKTYPE\GTYPE", "EqualVal", 77)
        self.lf.AddCondition("OP_AND", False, "ZK_OBERTYP_VERSCHIEDENE", "NotEqualVal", "*,*")

        if self.Visum.Net.Links.CountActive:

            # Erstelle einen DataFrame aus diesen Strecken
            atts_dict_withdtype = {"NO": int, "FROMNODENO": int, "TONODENO": int, "LENGTH": float, r"FROMNODE\CONCATENATE:INLINKS\NAME": str,
                                   r"TONODE\CONCATENATE:INLINKS\NAME": str}
            df_links = self.get_multi_netobj_atts_with_id(self.Visum.Net.Links, atts_dict_withdtype)
            # Straßennamen an Von und Nach-Knoten
            df_links['TFNAME'] = (df_links[r"FROMNODE\CONCATENATE:INLINKS\NAME"] + ',' + df_links[r"TONODE\CONCATENATE:INLINKS\NAME"]).apply(
                lambda a: set([x for x in a.split(',') if x != '']))
            # Duplikate entfernen (Spalte "NO" --> Hin/Rückrichtung):
            # df_links.drop_duplicates('NO', inplace=True)
            # Erstelle einen DiGraphen aus dem Dataframe mit möglichen U-Turns
            G = nx.from_pandas_edgelist(df_links, 'FROMNODENO', 'TONODENO', edge_attr=['VISUM_ID', 'NO', 'LENGTH', 'TFNAME'])
            # Ermittle welche Strecken miteinander verbunden sind (Start der Rampe bis Ende)
            connected_components = list(nx.connected_components(G))

            # create a list to store the results for each group
            grouped_links = []

            for component in connected_components:
                # create a subgraph for the current component
                subgraph = G.subgraph(component)
                # get the VISUM_IDs of all edges in the subgraph
                ids = [d['VISUM_ID'] for _, _, d in subgraph.edges(data=True)]
                # get the LINKNOs of all edges in the subgraph
                nos = [d['NO'] for _, _, d in subgraph.edges(data=True)]
                # get the sum of the lengths of all edges in the subgraph
                length_sum = sum([d['LENGTH'] for _, _, d in subgraph.edges(data=True)])
                # get unique f/t names
                ftnames = set().union(*[d['TFNAME'] for _, _, d in subgraph.edges(data=True)])
                # store the result in a dictionary
                grouped_links.append({'visum_ids': ids, 'linknos': nos, 'total_length': length_sum, 'ftnames': ftnames})

            # convert the dictionary to a dataframe
            grouped_links_df = pd.DataFrame(grouped_links)

            # Anzahl der Strecken in der Gruppe hinzufügen
            grouped_links_df['group_size'] = grouped_links_df['visum_ids'].apply(len)
            # Anzahl der vorkommenden Straßennamen hinzufügen
            grouped_links_df['ftname_size'] = grouped_links_df['ftnames'].apply(len)

            # Es handelt sich höchstwahrscheinlich um U-Turns, wenn der Straßenname sich nicht ändert und die Länge des gesamten Rampenstreckenzugs weniger als 50m beträgt
            uturns = grouped_links_df[(grouped_links_df['ftname_size'] == 1) & (grouped_links_df["total_length"] <= 0.05)].copy()
            uturns['ZK_IS_UTURN'] = 1
            linksnos2edit = uturns.explode('linknos')[['linknos', 'ZK_IS_UTURN']]
            linksnos2edit.set_index(pd.Index(linksnos2edit.linknos.to_list()), inplace=True)
            df_links_index = pd.DataFrame(self.Visum.Net.Links.GetMultiAttValues("NO", OnlyActive=False), columns=["VISUM_ID", "NO"]).astype({"VISUM_ID": int, "NO": int})
            df_links_index['ZK_IS_UTURN'] = df_links_index["NO"].map(linksnos2edit['ZK_IS_UTURN']).fillna(0)

            self.Visum.Net.Links.SetMultiAttValues('ZK_IS_UTURN', df_links_index[['VISUM_ID', 'ZK_IS_UTURN']].to_numpy())

    def postprocess_gtypes(self, roundabout_uda):
        """
        1. ZK_OBERTYP_UNGERICHTET was set for ramps already (comma separated list as ramps connect multiple gtypes. It must now be set for all other links (singular gtype)
        2. ZK_OBERTYP_STRECKE is introduced as a directed link uda containing the gtype of the link, here the gtype of the roundabouts is changed as well to be equal to ramps.
        :param roundabout_uda: boolean label to identify links in the network that are roundabouts
        """
        # Setze für Strecken die keine Rampen sind ZK_OBERTYP_VERSCHIEDENE
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "ZK_OBERTYP_UNGERICHTET", "GreaterVal", -1)
        self.edit_link_atts.SetAttValue("ONLYACTIVE", "1")
        self.edit_link_atts.SetAttValue("RESULTATTRNAME", "ZK_OBERTYP_VERSCHIEDENE")
        self.edit_link_atts.SetAttValue("FORMULA", 'NUMTOSTR([ZK_OBERTYP_UNGERICHTET],0)')
        self.Visum.Net.EditAttribute(self.edit_link_atts)
        self.lf.Init()

        # Benutze ein eigenes Attribut für den Obertyp des Streckentyps (gerichtet), das noch um die Info über U-Turns ergänzt wird
        self.add_uda(self.Visum.Net.Links, 'ZK_OBERTYP_STRECKE', value_type=1, def_val=99)
        self.edit_link_atts.SetAttValue("ONLYACTIVE", "0")
        self.edit_link_atts.SetAttValue("RESULTATTRNAME", "ZK_OBERTYP_STRECKE")
        self.edit_link_atts.SetAttValue("FORMULA", 'IF([ZK_IS_UTURN] & [LINKTYPE\GTYPE] != 99, 88, [LINKTYPE\GTYPE])')
        self.Visum.Net.EditAttribute(self.edit_link_atts)

        # Nachkorrektur von Kreisverkehren (ggf.), da Info nicht in [LINKTYPE\GTYPE]
        self.edit_link_atts.SetAttValue("RESULTATTRNAME", "ZK_OBERTYP_STRECKE")
        if roundabout_uda is not None and self.Visum.Net.Links.AttrExists(roundabout_uda):
            if self.Visum.Net.Links.AttrExists('OSM_CONSTRUCTION'):
                self.edit_link_atts.SetAttValue("FORMULA", f'IF([{roundabout_uda}]=1 & [OSM_CONSTRUCTION]=0, 77, [ZK_OBERTYP_STRECKE])')
            else:
                self.edit_link_atts.SetAttValue("FORMULA", f'IF([{roundabout_uda}]=1, 77, [ZK_OBERTYP_STRECKE])')
            self.Visum.Net.EditAttribute(self.edit_link_atts)


    def determine_nodetype(self):
        """ Ermittelt aus den vorab belegten Obertypen der Strecken Knotentypen

        :param:
        "CONCATENATEACTIVE:INLINKS\ZK_OBERTYP_VERSCHIEDENE": ungerichtet,
        "CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_UNGERICHTET": ungerichtet,
        "CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_STRECKE": gerichtet(in),
        "CONCATENATEACTIVE: OUTLINKS\ZK_OBERTYP_STRECKE": gerichtet(out)

        :return: Ergebnis im default Visum-Knotenattribut TYPENO
        """
        logging.info('Identification of intersection nodes and determination of node types based on the adjacent links (global link types).')
        # Ergebnis-Attribut
        self.add_uda(self.Visum.Net.Nodes, 'ZK_TYP_DETAIL', value_type=5, def_val='')
        # Filtern aller Strecken mit ZK_OBERTYP_UNGERICHTET kleiner 5
        self.Visum.Filters.InitAll()
        self.lf.AddCondition("OP_NONE", False, "ZK_OBERTYP_UNGERICHTET", "LessEqualVal", 5)

        # Egal ob OUTLINKS oder INLINKS wegen "undirected" Attributen
        nodelist_colnames_dtypes = {
            "NO": int,
            "CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_VERSCHIEDENE": str,
            # "CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_UNGERICHTET": str,
            "CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_STRECKE": str,
            "CONCATENATEACTIVE: OUTLINKS\ZK_OBERTYP_STRECKE": str
        }

        # Knotenliste in Visum anzeigen zu Debug-Zwecken (Nachvollziehbarkeit)
        nl_dbg = self.Visum.Lists.CreateNodeList
        nl_dbg.Show()
        nl_dbg.SetObjects(True)
        nl_dbg.AddColumn('TYPENO')
        nl_dbg.AddColumn('ZK_TYP_DETAIL')
        for att in nodelist_colnames_dtypes.keys():
            nl_dbg.AddColumn(att)

        # Knotenliste abrufen als DataFrame
        df_nodes_linksIO = self.get_multi_netobj_atts_with_id(self.Visum.Net.Nodes, nodelist_colnames_dtypes)

        # Zeilen ohne Einträge entfernen --> i.d.R. lose Enden ohne Zuordnung
        # Function to check if a string contains only comma-separated integers
        def is_comma_separated_integers(s):
            # Split the string by commas
            parts = s.split(',')
            # Check if all parts are digits
            return all(part.isdigit() for part in parts)

        # Apply the function and filter the DataFrame
        df_nodes_linksIO = df_nodes_linksIO[df_nodes_linksIO["CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_VERSCHIEDENE"].apply(is_comma_separated_integers)]

        # überarbeiten der Spalten mit Typen ([1:]) in sortierte listen
        str_cols_rework = list(nodelist_colnames_dtypes.keys())[1:]
        # Apply the transformation
        df_nodes_linksIO[str_cols_rework] = df_nodes_linksIO[str_cols_rework].apply(lambda col: col.map(lambda x: sorted([int(y) for y in x.split(',')])))
        # Erstellen einer Spalte InOut (IO)
        df_nodes_linksIO["IO"] = df_nodes_linksIO.apply(lambda x: sorted(x["CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_STRECKE"] + x["CONCATENATEACTIVE: OUTLINKS\ZK_OBERTYP_STRECKE"]),
                                                        axis=1)

        def process_column_nodetype_fromconcat(x):
            """Leite aus der Verkettung von Obertypen an In und Outlinks einen Knotentypen ab. Der Knotentyp bildet sich in der Regel aus dem höchsten X und zweithöchsten Y vorkommenden Obertyp X-Y.
            Eine Ausnahme sind Knoten mit nur einem vorkommenden Obertypen. Sie bilden einen X-X Typ
            Eine weitere Ausnahme sind Knoten, die mehr als 4 Ein/Ausgangsstrecken des hochrangigen Typs haben. Diese sind auch gleichrangige X-X Knoten
            """
            # String der verschiedenen ungerichtete vorkommenden Obertypen in eine Integer Liste umwandeln
            nums_dist = x["CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_VERSCHIEDENE"]
            # Ein- und Ausgehende Strecken die keine Rampen sind (77) und auch nicht gesperrt (99) aus den gerichteten Strecken ermitteln
            # Dies spielt eine Rolle bei der Bestimmung, ob ein 4-armiger Knoten der bspw den Obertyp 2 und 3 verbindet ein 2-2er oder 2-3er Knoten ist
            nums_dist_dir = [i for i in x["IO"] if i < 77]
            # nums_undir = [i for i in edit_link_atts["CONCATENATEACTIVE: INLINKS\ZK_OBERTYP_UNGERICHTET"] if i > -1]
            # umwandlung in ein set --> entfernen von Duplikaten
            set_nums = set(nums_dist)
            # Umwandlung in eine sortierte Liste (Reihenfolge wichtig!)
            nums_unq = sorted(list(set_nums))

            # Mehr als ein vorkommender Typ
            if len(nums_unq) > 1:
                # Initiales Ergebnis 'höchstes-zweithöchstes'
                result = f'{nums_unq[0]}-{nums_unq[1]}'
                # Editieren falls in der gerichteten Betrachtung mehr als vier Vorkommnisse des geringsten Obertyps (min) sind.
                # In diesem Fall ist der Knoten ein gleichrangiger Knoten X-X
                if len(nums_dist_dir) > 0:
                    if nums_dist_dir.count(min(nums_dist_dir)) > 4:
                        result = f'{min(nums_dist_dir)}-{min(nums_dist_dir)}'
                return result
            # Nur ein vorkommender Typ
            else:
                return f'{nums_unq[0]}-{nums_unq[0]}'

        # Bestimme den Knotentyp_Detail im Stil (X-Y, bzw X-X)
        df_nodes_linksIO["Knotentyp_Detail"] = df_nodes_linksIO.apply(process_column_nodetype_fromconcat, axis=1)
        # Aggregiere nach dem niedrigeren Rang
        df_nodes_linksIO["Knotentyp_Aggregiert"] = df_nodes_linksIO["Knotentyp_Detail"].apply(lambda row: int(row.split('-')[1]))
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_TYP_DETAIL", df_nodes_linksIO[["VISUM_ID", "Knotentyp_Detail"]].to_numpy())
        self.Visum.Net.Nodes.SetMultiAttValues("TYPENO", df_nodes_linksIO[["VISUM_ID", "Knotentyp_Aggregiert"]].to_numpy())

        # TYPENO initialisieren mit 99
        self.Visum.Net.Nodes.SetAllAttValues("TYPENO", 99, OnlyActive=False)

        # Identify nodes belonging to intersections following a set of criteria, considering the topology and hierarchy of adjacent links
        def set_nodetypes(row):
            # IO also counts blocked links and thus allows to determine an undirected degree
            links_io_incl_blocked = row["IO"]
            links_io_excl_blocked = [i for i in links_io_incl_blocked if i < 88]

            # Undirected degree (counting all links, including blocked ones divided by 2)
            undirected_degree = len(links_io_incl_blocked)/2

            # Total degree (sum of incoming and outgoing links excluding blocked ones)
            total_degree = len(links_io_excl_blocked)

            # Apply the criteria to identify if the node ibelongs to an intersection
            criterion1 = total_degree >= 5  # total degree equal to or greater than 5
            criterion2 = undirected_degree >= 4  # undirected degree equal to or greater than 4
            criterion3 = undirected_degree == 3 and (77 in links_io_incl_blocked)  # Check if undirected degree is 3 and if one of the adjacent links is a ramp or roundabout
            # TODO: Check if ramps should be ignored in crit 4. Ramps should be considered in
            criterion4 = (len(set(i for i in links_io_excl_blocked if i < 77)) > 1)  # Check if there's a change in the hierarchy among adjacent links (excluding ramps)

            # If any of the criteria is met return the nodetype meaning that this is actually an intersection node
            if any([criterion1, criterion2, criterion3, criterion4]):
                return row["Knotentyp_Aggregiert"]
            # else return 99 meaning it is not an intersection node
            else:
                return 99
        df_nodes_linksIO["Knotentyp_Aggregiert"] = df_nodes_linksIO.apply(set_nodetypes, axis=1)
        # Wenn TYPENO = 99 sollte dann ZK_TYP_DETAIL auf einen ungültigen Wert (N-N) setzen
        df_nodes_linksIO["Knotentyp_Detail"] = df_nodes_linksIO.apply(lambda row: row["Knotentyp_Detail"] if row["Knotentyp_Aggregiert"] < 99 else 'N-N', axis=1)

        # store in self.intersection_nodes_df
        self.intersection_nodes_df = df_nodes_linksIO.copy()

        # Write result back to Visum
        self.Visum.Net.Nodes.SetMultiAttValues("TYPENO", df_nodes_linksIO[["VISUM_ID", "Knotentyp_Aggregiert"]].to_numpy())
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_TYP_DETAIL", df_nodes_linksIO[["VISUM_ID", "Knotentyp_Detail"]].to_numpy())

    def set_open_end_nodes(self, spatial_selection_territory_no: int):
        logging.info('Set node types for open ends.')
        # Knoten und Strecken entsprechend filtern
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "ZK_OBERTYP_UNGERICHTET", "EqualVal", 1)
        self.nf.Init()
        self.nf.AddCondition("OP_NONE", False, "COUNTACTIVE:INLINKS", "EqualVal", 1)
        # Randknoten sollen hierbei nicht mitaufgenommen werden
        if spatial_selection_territory_no is not None:
            self.nf.AddCondition("OP_AND", False, "ZK_IST_RANDKNOTEN", "EqualVal", 0)

        self.Visum.Net.Nodes.SetAllAttValues("TYPENO", 0, OnlyActive=True)
        self.Visum.Net.Nodes.SetAllAttValues("ZK_TYP_DETAIL", '1-1', OnlyActive=True)
        # In der Regel ergeben sich Cluster aus 2+ Knoten, die nicht gelöscht werden sollen
        self.Visum.Filters.InitAll()

        # TODO: Update self.intersection_nodes_df accordingly...

    def create_cluster_actnodes(self, buffer, uda_name, nodecounter):
        """
        intersect and concat NO of active nodes with specified puffer, create clusters with unique numbers and write back to nodes
        returns a dataframe with main_nodes (sets of nodes and ZK_CLUSTER_ID)
        :param buffer: Beidseitiger Puffer [m] in dem zusammengehörige Knoten liegen dürfen
        :return:
        """

        # Visum-Operation Verschneiden mit Puffer
        intersect_att = self.Visum.Net.CreateIntersectAttributePara()
        intersect_att.SetAttValue("SOURCENETOBJECTTYPE", "NODE")
        intersect_att.SetAttValue("SOURCENETOBJECTINCLUDESUBCATEGORIES", "0")
        intersect_att.SetAttValue("SOURCEONLYACTIVE", "1")
        intersect_att.SetAttValue("SOURCEBUFFERSIZE", f"{buffer}m")
        intersect_att.SetAttValue("DESTNETOBJECTTYPE", "NODE")
        intersect_att.SetAttValue("DESTNETOBJECTINCLUDESUBCATEGORIES", "0")
        intersect_att.SetAttValue("DESTONLYACTIVE", "1")
        intersect_att.SetAttValue("DESTBUFFERSIZE", f"{buffer}m")
        intersect_att.SetAttValue("RANKATTRNAME", "SPECIALENTRY_EMPTY")
        intersect_att.SetAttValue("SMALLRANKIMPORTANT", "1")
        intersect_att.SetAttValue("SOURCEATTRNAME", "NO")
        intersect_att.SetAttValue("DESTATTRNAME", f"{uda_name}")
        intersect_att.SetAttValue("NUMERICOPERATION", "INTERSECTION_SUM")
        intersect_att.SetAttValue("STRINGOPERATION", "INTERSECTION_CONCATENATE")
        intersect_att.SetAttValue("ROUND", "0")
        intersect_att.SetAttValue("ADDVALUE", "0")
        intersect_att.SetAttValue("WEIGHTBYINTERSECTIONAREASHARE", "1")
        intersect_att.SetAttValue("CONCATMAXLEN", "999999")
        intersect_att.SetAttValue("CONCATSEPARATOR", ",")

        self.Visum.Net.IntersectAttributes(intersect_att)

        # Erstelle eine Liste der Knoten mit einer Spalte mit den Überlappungen aus dem Verschneiden Schritt
        # Attribute für GetAtt
        nodelist_colnames_dtypes = {
            "NO": int,
            "ZK_TYP_DETAIL": str,
            "ZK_NODES_INTERSECT": str,
            "XCOORD": float,
            "YCOORD": float,
        }
        # Als DataFrame einlesen
        df_nodes = self.get_multi_netobj_atts_with_id(self.Visum.Net.Nodes, nodelist_colnames_dtypes)

        # Umwandeln der kommaseparierten Überlappungen in sets
        node_sets_individual = [set(int(i) for i in x.split(',')) for x in df_nodes['ZK_NODES_INTERSECT']]

        # Gemeinsame Sets, die aktualisiert werden
        node_sets_joint = []

        # Iteriere über jedes individuelle Set, überprüfe in jeder Schleife ob durch das aktuelle set neue Knoten dazukommen oder bisher getrennt vorliegende sets vereint werden
        for s in node_sets_individual:
            # Finde alle Sets in node_sets_joint, die Überlappungen mit dem aktuellen individuellen Set haben
            overlaps = [r for r in node_sets_joint if not r.isdisjoint(s)]
            # Kombiniere das aktuelle individuelle Set mit allen überlappenden Sets
            temp_s = s | set().union(*overlaps)
            # Entferne die überlappenden Sets aus node_sets_joint
            node_sets_joint = [r for r in node_sets_joint if r not in overlaps]
            # Füge das aktualisierte Set zu node_sets_joint hinzu
            node_sets_joint.append(temp_s)

        # Das resultierende node_sets_joint enthält die vereinigten und aktualisierten Sets
        df_mainnodes = pd.DataFrame({"node_clusters": node_sets_joint})

        # Add a column to df_mainnodes to store the count of nodes in each cluster
        df_mainnodes["node_count"] = df_mainnodes["node_clusters"].apply(len)
        df_mainnodes['cluster_id'] = df_mainnodes.index + nodecounter

        # Iterate over the rows of df_mainnodes and update df_nodes accordingly
        for index, row in df_mainnodes.iterrows():
            node_cluster = row["node_clusters"]
            node_count = row["node_count"]
            cluster_id = row['cluster_id']

            # Update df_nodes with the node count for the corresponding nodes in the node cluster
            df_nodes.loc[df_nodes['NO'].isin(node_cluster), 'ZK_CLUSTER_ANZAHL_KNOTEN'] = node_count
            # Update df_nodes with the index of the cluster for the corresponding nodes in the node cluster
            df_nodes.loc[df_nodes['NO'].isin(node_cluster), 'ZK_CLUSTER_ID'] = cluster_id

        return df_nodes

    def cluster_by_nodetype(self, nodetype_buffer_dict):
        # Knoten-Attribut zum Verketten der verschnittenen anderen Knoten innerhalb der Puffer [Typ: langer Text (62)]
        self.add_uda(self.Visum.Net.Nodes, 'ZK_NODES_INTERSECT', value_type=62, def_val='')
        # Knoten-Attribute zum Festhalten der Cluster-ID und der Anzahl der Cluster im Knoten
        self.add_uda(self.Visum.Net.Nodes, 'ZK_CLUSTER_ID', value_type=1, def_val=0)
        self.add_uda(self.Visum.Net.Nodes, 'ZK_CLUSTER_ANZAHL_KNOTEN', value_type=1, def_val=0)

        # Erstelle Counter für neu einzufügende Netzobjekte
        # Zugangsknoten: an diese wird später angebunden. Sie aggregieren die gebildeten Cluster an der geometrischen Zentroide
        zugangsknoten_counter = self.new_object_number('Nodes')
        # # Zugangsstrecken verbinden die Zugangsknoten mit dem vorhandenen Netz
        # zugangsstrecken_counter = self.new_object_number('Links')

        # Erstelle einen DataFrame, in dem die einzufügenden neuen Zugangsknoten aufgenommen werden
        zugangsknoten_master = pd.DataFrame()

        # Arbeite die Knotentypen in knotentyp_puffer_dict der Reihenfolge nach ab
        for knotentyp, (puffer) in nodetype_buffer_dict.items():
            # Filter auf Knoten des entsprechenden Typs
            self.nf.Init()
            self.nf.AddCondition("OP_NONE", False, "ZK_TYP_DETAIL", "EqualVal", knotentyp)

            # Falls ein Puffer definiert ist (>0), dann clustere mit diesem Puffer
            if puffer > 0:
                # Prüfen, ob es Knoten dieser Art im Netz gibt
                if self.Visum.Net.Nodes.CountActive > 0:
                    logging.info(f'Clustering nodes of intersection node type: {knotentyp} with {puffer}m buffer')
                    # Cluster der aktiven Knoten mit angegebenen Puffer bilden
                    df_nodes = self.create_cluster_actnodes(puffer, 'ZK_NODES_INTERSECT', zugangsknoten_counter)
                    zugangsknoten_counter = df_nodes["ZK_CLUSTER_ID"].max() + 1
                    zugangsknoten_master = pd.concat([zugangsknoten_master, df_nodes])
                else:
                    logging.info(f'No nodes of intersection node type: {knotentyp} in the network.')
                    continue
            # Ist kein Puffer definiert (=0), dann werden alle Knoten ohne Clustern mit Puffer in zugangsknoten_master aufgenommen
            else:
                # Typ 5-5 ohne clustern aufnehmen
                logging.info(f'Inclusion of nodes of intersection node type: {knotentyp} (no clustering required).')
                nodelist_colnames_dtypes = {
                    "NO": int,
                    "ZK_TYP_DETAIL": str,
                    "ZK_NODES_INTERSECT": str,
                    "XCOORD": float,
                    "YCOORD": float,
                }
                # Als DataFrame einlesen und zum master hinzufügen
                df_nodes = self.get_multi_netobj_atts_with_id(self.Visum.Net.Nodes, nodelist_colnames_dtypes)
                df_nodes["ZK_CLUSTER_ID"] = df_nodes.index + zugangsknoten_counter
                zugangsknoten_counter = df_nodes["ZK_CLUSTER_ID"].max() + 1
                df_nodes["ZK_CLUSTER_ANZAHL_KNOTEN"] = 1
                zugangsknoten_master = pd.concat([zugangsknoten_master, df_nodes])

        # Attribute nach Visum übertragen
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_CLUSTER_ID", zugangsknoten_master[["VISUM_ID", "ZK_CLUSTER_ID"]].to_numpy())
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_CLUSTER_ANZAHL_KNOTEN", zugangsknoten_master[["VISUM_ID", "ZK_CLUSTER_ANZAHL_KNOTEN"]].to_numpy())

        zugangsknoten_master.set_index("VISUM_ID", drop=False, inplace=True)

        self.access_nodes_df = zugangsknoten_master.copy()

    def prepare_net_file(self):
        logging.info('Preparing results to fit with Visum object structure for netfile')
        zugangsknoten_master = self.access_nodes_df.copy()
        # Alle in zugangsknoten_master gesetzten/geänderten Werte zurück nach andi.Visum schreiben
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_CLUSTER_ID", zugangsknoten_master[["VISUM_ID", "ZK_CLUSTER_ID"]].to_numpy())
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_CLUSTER_ANZAHL_KNOTEN", zugangsknoten_master[["VISUM_ID", "ZK_CLUSTER_ANZAHL_KNOTEN"]].to_numpy())
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_TYP_DETAIL", zugangsknoten_master[["VISUM_ID", "ZK_TYP_DETAIL"]].to_numpy())

        # 3.1 Vorbereitung Knoten
        # Eine Typ-Nummer ergänzen, die nachher beim Anbinden verwendet werden kann.
        self.add_uda(self.Visum.Net.Nodes, uda_name='ZK_TYP', value_type=1, def_val=99,
                     comment='Typ-Nummer des Zugangsknotens. Bei der Nummer 99 handelt es sich um keinen Zugangsknoten.')

        # UDA vorbereiten in zugangsknoten_master
        zugangsknoten_master['ZK_TYP'] = zugangsknoten_master['ZK_TYP_DETAIL'].apply(lambda row: int(row.split('-')[1]))

        # Aggregiere über die ZK_CLUSTER_ID um eine Liste einzufügender neuer Knoten zu erhalten an die später angebunden werden soll
        dict_ranks = {key: int(key.replace('-', '')) for key in set(zugangsknoten_master["ZK_TYP_DETAIL"])}
        aggregation_dict = {
            "VISUM_ID": set,
            "NO": set,
            "XCOORD": 'mean',
            "YCOORD": 'mean',
            "ZK_TYP_DETAIL": lambda x: min(x, key=lambda y: dict_ranks[y]),
            "ZK_TYP": 'min'
        }
        zk_master_agg = zugangsknoten_master.groupby(["ZK_CLUSTER_ID"]).agg(aggregation_dict)
        zk_master_agg["ZK_CLUSTER_ID"] = zk_master_agg.index
        zk_master_agg["ZK_CLUSTER_ANZAHL_KNOTEN"] = zk_master_agg["VISUM_ID"].apply(len)

        # Wenn nach dem Clustern nur ein Knoten im Cluster befindlich ist, dann soll kein neuer eingefügt werden
        # Verwende hier den original OSM-Knoten für die entsprechende Attribuierung als Zugangsknoten (viele 5-5er)
        mask_1ercluster = (zk_master_agg.ZK_CLUSTER_ANZAHL_KNOTEN == 1)
        df_1ercluster = zk_master_agg[mask_1ercluster]
        self.Visum.Net.Nodes.SetMultiAttValues("ZK_TYP", df_1ercluster.explode("VISUM_ID")[["VISUM_ID", "ZK_TYP"]].values)

        # Die restlichen Cluster benötigen einen neu eingefügten Zugangsknoten. Bereite einen df so vor, dass er import tauglich wird
        df_newnodes = zk_master_agg[~mask_1ercluster][['ZK_CLUSTER_ID', 'ZK_TYP_DETAIL', 'XCOORD', 'YCOORD', 'ZK_TYP']].copy()
        df_newnodes.rename(columns={'ZK_CLUSTER_ID': 'NO'}, inplace=True)
        df_newnodes.index = df_newnodes.index.astype(int)

        # Store in instance object placeholder
        self.nodes4visum_df = df_newnodes.copy()

        # 3.2 Vorbereitung Strecken
        # Streckentyp 777 als default für eingefügte Strecken
        # Tsys, Kapazität, Numlanes wird über den Streckentyp automatisch gesetzt
        # Ein Wert von 20km/h bedeutet, dass eher kurze Strecken vom Zugangsknoten verwendet werden, um möglichst schnell an das existierede Netz zu kommen.
        # Abbieger werden auch automatisch gesetzt beim Einlesen der Net-Datei
        lt777 = self.Visum.Net.AddLinkType(777)
        lt777.SetAttValue("NAME", 'ZK_Hilfsstrecken')
        lt777.SetAttValue("V0PRT", 20)
        lt777.SetAttValue("NUMLANES", 1)
        lt777.SetAttValue("CAPPRT", 9999)

        # Strecken die den neuen Zugangsknoten mit dem vorhandenen Netz verbinden aus den Knotenlisten ableiten
        zk_master_agg['Links'] = zk_master_agg.apply(
            lambda row: [(row['ZK_CLUSTER_ID'], i) for i in row['NO']] + [(i, row['ZK_CLUSTER_ID']) for i in row['NO']], axis=1)
        df_newlinks = pd.DataFrame({"Links": zk_master_agg.Links.explode()})
        df_newlinks[['FromNodeNo', 'ToNodeNo']] = pd.DataFrame(df_newlinks['Links'].tolist(), index=df_newlinks.index)

        # Sortiere die Links so, dass Hin und Rückrichtung beieinanderstehen (zb nach original Knoten) wie in self.Visum default Listenlayout (lesbarkeit)
        df_newlinks["FromToNode_Sort"] = df_newlinks[['FromNodeNo', 'ToNodeNo']].min(axis=1)
        df_newlinks.sort_values(by=['FromToNode_Sort', 'FromNodeNo'], inplace=True)
        df_newlinks.drop(["Links", "FromToNode_Sort"], axis=1, inplace=True)
        df_newlinks.reset_index(drop=True, inplace=True)
        # Zugangsstrecken verbinden die Zugangsknoten mit dem vorhandenen Netz
        zugangsstrecken_counter = self.new_object_number('Links')
        # Ergänze Spalten für Nummer, Typ und TSys
        df_newlinks.insert(0, "No", df_newlinks.index.repeat(2)[:len(df_newlinks)] + zugangsstrecken_counter)
        df_newlinks['TypeNo'] = 777
        df_newlinks['ZK_OBERTYP_STRECKE'] = 777

        # Store in instance object placeholder
        self.links4visum_df = df_newlinks.copy()

    def create_net_file(self, netfile_path):
        logging.info('Writing net-file and loading to Visum')
        file = open(netfile_path, "w")
        nodes_df = self.nodes4visum_df.copy()
        links_df = self.links4visum_df.copy()

        # header
        file.write('''$VISION
* Universität Stuttgart Fakultät 2 Bau+Umweltingenieurwissenschaften Stuttgart
* 05.07.23
* 
* Tabelle: Versionsblock
* 
$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT
12.000;Net;ENG;KM
''')

        # nodes-header
        file.write(f'''
*
*Tabelle: Knoten
*
$NODE:{";".join(nodes_df.columns).upper()}
''')
        # nodes
        nodes_df.to_csv(file, header=False, sep=";", index=False, lineterminator='\n')
        # links-header
        file.write(f'''
*
*Tabelle: Strecken
*
$LINK:{";".join(links_df.columns).upper()}
''')
        # links
        links_df.to_csv(file, header=False, sep=";", index=False, lineterminator='\n')

        zone_df = nodes_df[['NO', 'XCOORD', 'YCOORD']].copy()
        # zone_df = nodes_df.drop(['ZK_TYP_DETAIL', 'ZK_TYP'], axis=1)
        zone_df["NODENO"] = zone_df["NO"]
        new_zone_no = self.new_object_number('Zones')
        zone_df["NO"] = zone_df.index + new_zone_no
        zone_df["TYPENO"] = 99
        # zone-header
        file.write(f'''
*
*Tabelle: Bezirke
*
$ZONE:{";".join(zone_df.drop("NODENO", axis=1).columns).upper()}
''')
        # zones
        zone_df.drop("NODENO", axis=1).to_csv(file, header=False, sep=";", index=False, lineterminator='\n')

        conn_df = pd.DataFrame({'ZONENO': zone_df['NO'], 'NODENO': zone_df['NODENO']})
        conn_df = pd.concat([conn_df, conn_df], axis=0)
        conn_df.sort_values('ZONENO', inplace=True)
        conn_df.reset_index(drop=True, inplace=True)
        conn_df['DIRECTION'] = conn_df.apply(lambda x: 'O' if x.name % 2 == 0 else 'D', axis=1)
        # Öffnen für alle TSYS die es in Visum gibt
        conn_df['TSYSSET'] = ','.join([tsys.AttValue("CODE") for tsys in self.Visum.Net.TSystems.GetAll])

        # connector-header
        file.write(f'''
*
*Tabelle: Anbindungen
*
$CONNECTOR:{";".join(conn_df.columns).upper()}
''')
        # conns
        conn_df.to_csv(file, header=False, sep=";", index=False, lineterminator='\n')

        file.close()
        # Open in Visum
        self.Visum.LoadNet(netfile_path, ReadAdditive=True)

    def postprocess_helper_links(self, Visum_helper_files_path):
        logging.info('Deactivating non-required helper links through simple assignment.')
        # 1er Matrix mit einfacher Widfkt...
        # --> neu eingefügte Strecken mit hohem Widerstand
        # --> Rampen (Typ "_linK", ZK_OBERTYP_UNGERICHTET) ebenfalls hoher Widerstand
        # Möglichst sollen die neue Strecken von den neuen Knoten direkt ans Netz und nicht erst auf die Rampen.

        # Testen der Anbindungsknoten + ausdünnen der Anbindungsknoten_Hilfstrecken
        # Alle alten Matrizen löschen
        self.Visum.Net.Matrices.RemoveAll()
        # Eine initiale Matrix erzeugen (1er Matrix)
        # 1er Matrix (100) anlegen um Anbindungen zu testen
        matrix_1er = self.AddMat(ObjId=1, Mat_Code="1erUmlegung_TestAnbindungsknoten", Mat_Name="1erUmlegung_TestAnbindungsknoten", Objecttype=2,
                                 Matrixtype=3, Formula="")

        # DSeg einfügen zum Testen der Anbindungen über 1erMatrix
        # DSegs Pkw (könnte später noch geändert werden zu Fuss/Rad)
        self.Visum.Net.DemandSegments.RemoveAll()
        dseg_1erumlegung_name = "1erUmlegung_TestAnbindungsknoten"
        # # Versuche ein DSeg mit Namen/Code P, PKW, Car, C zu finden
        no_pkw_mode = True
        for m in self.Visum.Net.Modes.GetAll:
            if {m.AttValue("NAME"), m.AttValue("CODE")} & {"Pkw", "P", "Car", "C"}:
                # print(f'Modus für Pkw: {m.AttValue("NAME"), m.AttValue("CODE")}')
                dseg_1erumlegung = self.Visum.Net.AddDemandSegment(dseg_1erumlegung_name, m.AttValue("CODE"))
                dseg_1erumlegung.SetAttValue("NAME", dseg_1erumlegung_name)
                dseg_1erumlegung.getDemandDescription().SetAttValue("Matrix", 'Matrix([CODE]="' + dseg_1erumlegung_name + '")')
                no_pkw_mode = False
        if no_pkw_mode:
            raise TypeError('Es konnte kein Modus für Pkw ("Pkw", "P", "Car", "C") in self.Visum gefunden werden.')

        # dseg_1erumlegung = self.Visum.Net.AddDemandSegment(dseg_1erumlegung_name, code_pkw_mode)
        # dseg_1erumlegung.SetAttValue("NAME", dseg_1erumlegung_name)
        # dseg_1erumlegung.getDemandDescription().SetAttValue("Matrix", 'Matrix([CODE]="' + dseg_1erumlegung_name + '")')

        # 1er Nachfragematrix erzeugen
        matrix_1er.SetValuesToResultOfFormula("If(FROM[TYPENO]=99 & TO[TYPENO]=99, 1,0)")
        matrix_1er.SetAttValue("DSEGCODE", dseg_1erumlegung_name)

        # Verfahrensablauf mit Umlegung einlesen und ausführen
        umlegung1er_procedure_path = Visum_helper_files_path / '20230814_1erUmlegung_TestAnbindungsknoten.xml'
        self.Visum.Procedures.Open(umlegung1er_procedure_path)
        self.Visum.Procedures.Execute()

        # Alle Strecken die in dieser Umlegung nicht verwendet werden bekommen einen anderen Streckentyp 776 und werden für später gesperrt
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "TYPENO", "ContainedIn", "777")
        self.lf.AddCondition("OP_AND", False, "VOLVEHPRT(AP)", "EqualVal", 0)

        lt776 = self.Visum.Net.AddLinkType(776)
        lt776.SetAttValue("NAME", 'ZK_Hilfsstrecken_Überflüssig')
        lt776.SetAttValue("V0PRT", 0)
        lt776.SetAttValue("NUMLANES", 0)
        lt776.SetAttValue("CAPPRT", 0)

        # Überschreibe die Setzungen aus der Net-Datei
        self.Visum.Net.Links.SetAllAttValues("TSYSSET", "", OnlyActive=True)
        self.Visum.Net.Links.SetAllAttValues("ZK_OBERTYP_STRECKE", 776, OnlyActive=True)
        self.Visum.Net.Links.SetAllAttValues("TYPENO", 776, OnlyActive=True)

        self.Visum.Filters.InitAll()

    def save_results(self, working_dir, version_filename, Visum_helper_files_path, write_debug_file=True):
        # prepare filenames
        orig_ver_path = working_dir / version_filename
        # debug file
        debug_ver_file = working_dir / f'{orig_ver_path.stem}_AccessNodes_DebugFile.ver'
        # final file
        final_ver_file = working_dir / f'{orig_ver_path.stem}_AccessNodes.ver'

        if write_debug_file:
            debug_ver_file.parent.mkdir(parents=False, exist_ok=True)
            self.nf.AddCondition("OP_NONE", False, "ZK_TYP", "LessVal", 99)
            self.Visum.Graphic.StopDrawing = False
            self.Visum.Graphic.ShowMaximized()
            self.Visum.Net.GraphicParameters.Open(Visum_helper_files_path / '20241201_AccessNodeType_Detail.gpa')
            self.Visum.Graphic.DisplayEntireNetwork()
            self.Visum.SaveVersion(debug_ver_file)
            logging.debug(f"Debug-File saved to: {debug_ver_file}")

        # Löschen der eingefügten Elemente die im Weiteren nicht benötigt werden
        self.Visum.Net.Matrices.RemoveAll()
        # Bezirke löschen (Anbindungen werden automatisch gelöscht, beim Bezirke löschen)
        zf = self.Visum.Filters.ZoneFilter()
        zf.Init()
        zf.AddCondition("OP_NONE", False, "TYPENO", "ContainedIn", "99")
        self.Visum.Net.Zones.RemoveAll(OnlyActive=True)

        # Löschen der weiterhin nicht benötigten udas
        uda_dict_del = {
            'LINK_IDENTIFIER': self.Visum.Net.Links,
            # 'ZK_IS_UTURN': self.Visum.Net.Links,
            'ZK_OBERTYP_UNGERICHTET': self.Visum.Net.Links,
            'ZK_OBERTYP_VERSCHIEDENE': self.Visum.Net.Links,
            # 'ZK_OBERTYP_STRECKE': self.Visum.Net.Links,
            'ZK_ÜBERGABEKNOTEN': self.Visum.Net.Nodes,
            # 'ZK_TYP_DETAIL': self.Visum.Net.Nodes,
            # "name_cluster_id": self.Visum.Net.Nodes,
            'ZK_IST_RANDKNOTEN': self.Visum.Net.Nodes,
            'ZK_NODES_INTERSECT': self.Visum.Net.Nodes,
            # 'ZK_CLUSTER_ID': self.Visum.Net.Nodes,
            # 'ZK_CLUSTER_ANZAHL_KNOTEN': self.Visum.Net.Nodes,
            # 'ZK_TYP': self.Visum.Net.Nodes
        }

        for uda, netobj in uda_dict_del.items():
            try:
                netobj.DeleteUserDefinedAttribute(uda)
            except:
                logging.debug(f'{uda} might not exist')

        # Set U_TURNs to ANBINDUNGSKNOTEN_OBERTYPSTRECKE 78
        # TODO: Implement this upon creation before 'bestimme_rin_nodetype()'...
        #  this is dangerous because other parts of the code might rely on it...
        self.lf.Init()
        self.lf.AddCondition("OP_NONE", False, "ZK_IS_UTURN", "EqualVal", 1)
        self.Visum.Net.Links.SetAllAttValues("ZK_OBERTYP_STRECKE", newValue=78, OnlyActive=True)
        self.lf.Init()

        # Originale Obertypen wieder einlesen
        self.Visum.IO.LoadAttributeFile(self.gtypes_original_att)

        self.Visum.Filters.InitAll()
        self.nf.AddCondition("OP_NONE", False, "ZK_TYP", "LessVal", 99)
        self.Visum.Graphic.StopDrawing = False
        self.Visum.Graphic.ShowMaximized()
        self.Visum.Net.GraphicParameters.Open(Visum_helper_files_path / '20241201_AccessNodeType_Detail.gpa')
        self.Visum.Graphic.DisplayEntireNetwork()
        self.Visum.SaveVersion(final_ver_file)
        logging.debug(f"Final Version saved to: {final_ver_file}")
        logging.debug((f"Total number of derived access nodes: {self.Visum.Net.Nodes.CountActive}"))

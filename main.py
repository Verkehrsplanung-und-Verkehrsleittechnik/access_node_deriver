from pathlib import Path
import pandas as pd
import logging
from helpers.logger_setup import setup_logger
from helpers.access_node_derivation import AccessNodeDeriver


if __name__ == '__main__':
    # Path to the folder with Visum version and for log files
    working_dir = Path('path/to/your/working/directory')

    # LOGGING
    setup_logger(log_file=working_dir / 'zugangsknoten_kfz.log', level=logging.INFO)

    # Filename of the Visum version
    ver_filename = 'visum_version.ver'
    # Path to Visum-related files needed (main.py file path/visum_helper_files)
    Visum_helper_files_path = Path(__file__).parent / 'visum_helper_files'

    # Settings via toggler
    set_gtypes_osm = True  # Set global link types based on OSM names (key: highway)? Otherwise, manual setting is required.
    spatial_selection_territory_no = 1  # Restrict analysis to a specific territory in Visum? Specify the number, otherwise None.
    # Specification of the link UDA indicating whether a link is a roundabout (Boolean) for further consideration. If not present: None
    roundabout_attribute = "OSM_ROUNDABOUT"

    logging.info(f'DERIVING ACCESS NODES FOR MOTORIZED VEHICLE TRAFFIC IN {ver_filename}')
    # Create AccessNodeDeriver Instance (andi) and initialize
    andi = AccessNodeDeriver(working_dir, ver_filename)

    # --------------------------------------------------
    # 1. NETWORK PREPARATION
    logging.info(f'Preparation of the network (global link types, node types): {ver_filename}')
    # 1.1 Global link types (gtypes) in Visum should be set to match the Functional Road Class
    # Specification: Values 1 - 5 for drivable roads (1 = major, 5 = residential street), 77 for ramps if known, 99 for others

    # Preparation of the network: Setting global link types
    andi.prepare_net(set_gtypes_osm, working_dir, Visum_helper_files_path)

    # 1.2 Restrict the area e.g., 30 km radius around the planning area
    if spatial_selection_territory_no is not None:
        andi.spatial_selection(spatial_selection_territory_no)


    # 1.3 Preprocessing of ramps, roundabouts, U-turns
    logging.info('Processing of links belonging to special infrastructures (ramps, roundabouts, u-turns).')

    # 1.3.1 Revision of roundabouts: Same processing as ramps. Set the same global link type (77) for roundabouts for this purpose
    if roundabout_attribute is not None:
        andi.preprocess_roundabouts(roundabout_attribute)

    # 1.3.2 Process ramps (and roundabouts)
    # At each ramp, all global types of the links they connect should be present
    # New link attribute 'ZK_OBERTYP_VERSCHIEDENE' for this purpose
    andi.add_uda(andi.Visum.Net.Links, 'ZK_OBERTYP_VERSCHIEDENE', value_type=5, def_val='')
    # Conduct ramp analysis
    andi.ramp_analysis()

    # 1.3.3 Identify U-turns
    # U-turns are often mapped as ramps in OSM. In determine_nodetype(), they would be identified as intersections through criterion3.
    # The following function sets their global type to 88 and thus excludes them from criterion3.
    andi.identify_uturns()

    # 1.3.4 Final editing of link attributes
    andi.postprocess_gtypes(roundabout_attribute)

    # --------------------------------------------------
    # 2. IDENTIFY INTERSECTION_NODES AND BUILD CLUSTERS TO DERIVE ACCESS NODES

    # 2.1 Identification of intersection nodes and determination of node types based on the adjacent links (global link types)
    andi.determine_nodetype()
    # NOTE: The result of the identification of intersection nodes can be examined in andi.intersection_nodes_df (can be used for further analysis or debugging).

    # 2.1.1 Processing of open ends of type 1 to 1-1
    andi.set_open_end_nodes(spatial_selection_territory_no)

    # 2.2 Create clusters from intersection nodes of the same node type to derive access nodes

    # Dictionary that specifies buffer distances for each possible intersection node type
    # Within the buffer distance, nodes of the same type are searched and clustered
    # The buffer is bidirectional, so for example two 0-0 nodes can be found up to 2000m apart
    nodetype_buffer_dict = {
        '1-1': 1000.0,
        '1-2': 500.0,
        '1-3': 250.0,
        '1-4': 250.0,
        '1-5': 250,
        '2-2': 250.0,
        '2-3': 125.0,
        '2-4': 125.0,
        '2-5': 12.5,
        '3-3': 150.0,
        '3-4': 50.0,
        '3-5': 12.5,
        '4-4': 50.0,
        '4-5': 12.5,
        '5-5': 0.0
    }

    andi.cluster_by_nodetype(nodetype_buffer_dict)
    # NOTE: The result are the access node. They can be examined in andi.access_nodes_df (can be used for further analysis or debugging).

    # -------------------------------------------------
    # 3. PREPARE NET-FILE AND IMPORT
    logging.info(f'Preparing and importing net-file containing the access nodes for {ver_filename}.')

    # 3.1 Preparation of the results for Visum
    andi.prepare_net_file()
    # NOTE: The result are the access nodes and helper links connecting the access nodes to all intersection nodes in each cluster.
    # They can be examined in andi.nodes4visum_df and andi.links4visum_df(can be used for further analysis or debugging).

    # 3.2 Create and load net-file to Visum
    netfile_path = working_dir / 'net' / f"{ver_filename}_helpers_connector_nodes.net"
    netfile_path.parent.mkdir(parents=False, exist_ok=True)
    andi.create_net_file(netfile_path)

    # 3.3 Identify and deactivate unnecessary helper links
    andi.postprocess_helper_links(Visum_helper_files_path)

    # 3.4 Save results to new version file(s)
    andi.save_results(working_dir, ver_filename, Visum_helper_files_path, write_debug_file=True)
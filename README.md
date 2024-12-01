# Access Node Deriver for Visum

This Python script derives access nodes for motorized vehicle traffic using a Visum network file.
It processes network data, identifies  specific nodes, and prepares a `.net` file for Visum import, enabling further use in Visum such as generating connectors

## Key Features
- **Network Preparation**: Sets global link types, processes ramps, roundabouts, and U-turns.
- **Intersection Detection**: Identifies nodes belonging to intersections (intersection nodes) based on topology and hierarchy of adjacent roads and assigns a node type.
- **Node Clustering**: Groups intersection nodes by node type using buffer distances.
- **Visum Integration**: Exports nodes and helper links to a `.net` file for import into Visum.

## How to Use
1. Make sure to set the global link types according to the following specification: Values 1(major) - 5(local) for drivable roads to match the functional road class
    - The script is optimized for OpenStreetMap networks with link hierarchies from the key 'highway'. The global link type can be set automatically (set_gtypes_osm()). Netwrok graphs from other sources can be used but they need a similar hierarchy in the links and the global link type must be set manually.
      - 1: motorway  
      - 2: trunk, primary  
      - 3: secondary  
      - 4: tertiary  
      - 5: unclassified, residential, road, living_street  
      - 6: service

2. Run the script to process the network and generate results.
3. Results are imported into Visum by the script in a new version file for further analysis.

## Requirements
- PTV Visum software
- Python 3.x with required dependencies

For more details, refer to the comments in the script.

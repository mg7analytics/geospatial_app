import streamlit as st
import geopandas as gpd
import pandas as pd
from shapely.wkt import loads
from shapely.geometry import Polygon, MultiPolygon
import io
from openpyxl import Workbook

# Custom CSS for styling
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F5F5F5;
        font-family: 'Roboto', sans-serif;
    }
    .header {
        background-color: #1976D2;
        color: white;
        padding: 10px 20px;
        border-radius: 5px;
        text-align: center;
        margin-bottom: 20px;
    }
    .header h1 {
        margin: 0;
        font-size: 28px;
    }
    .stProgress > div > div > div {
        background-color: #2E7D32;
        background: linear-gradient(to right, #2E7D32, #66BB6A);
        border-radius: 5px;
    }
    .stButton>button {
        background-color: #1976D2;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #1565C0;
        transform: scale(1.05);
    }
    .stFileUploader label {
        color: #424242;
        font-size: 16px;
        font-weight: bold;
    }
    .expander {
        border: 1px solid #B0BEC5;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Branded header
st.markdown('<div class="header"><h1>Geospatial Data Preprocessing App</h1></div>', unsafe_allow_html=True)

# Language selection
language = st.selectbox("Choose language / Choisir la langue", ["English", "French"])

# Instructions in English and French
instructions_en = """
### Instructions
#### File Requirements
- Upload a file in KML, Excel (.xlsx), or GeoJSON format.
- The file must contain:
  - A column named `wkt_geom` with Polygon or MultiPolygon geometries in WKT format.
  - A column named `Plantation Code` for attribute analysis.

#### Output Tabs in the Excel Report
- **all_data**: All loaded data (with count in parentheses).
- **geo_unique**: Unique geometries (with count).
- **geo_duplicates**: Duplicate geometries (with count).
- **attr_duplicates**: Duplicate Plantation Codes (with count).
- **attr_unique**: Unique Plantation Codes (with count).
- **num_point**: Polygons with fewer than 12 points (with count).
- **centroid**: Polygons whose centroids are outside their boundaries (with count).
- **ovlp15**: Polygons from `geo_unique` overlapping by more than 15%, including:
  - Plantation Code.
  - Area in hectares (3 decimals).
  - Overall overlap percentage (2 decimals).
  - Overlap percentage of each polygon relative to the other (2 decimals) (with count).
- **valid**: Filtered data starting from `geo_unique`, with the following conditions:
  - Excludes duplicate attributes.
  - Removes overlapping polygons (>15%) based on area (if two polygons have the same area, one is removed).
  - Ensures no duplicates in geometries or attributes remain after all filters.
  - Includes:
    - `area_ha` (3 decimals, 1 ha = 10,000 m²).
    - `longitude` (8 decimals).
    - `latitude` (8 decimals) (with count).
"""

instructions_fr = """
### Instructions
#### Exigences du fichier
- Téléchargez un fichier au format KML, Excel (.xlsx) ou GeoJSON.
- Le fichier doit contenir :
  - Une colonne nommée `wkt_geom` avec des géométries de type Polygon ou MultiPolygon au format WKT.
  - Une colonne nommée `Plantation Code` pour l'analyse des attributs.

#### Onglets de sortie dans le rapport Excel
- **all_data** : Toutes les données chargées (avec le nombre entre parenthèses).
- **geo_unique** : Géométries uniques (avec le nombre).
- **geo_duplicates** : Géométries dupliquées (avec le nombre).
- **attr_duplicates** : Codes de plantation dupliqués (avec le nombre).
- **attr_unique** : Codes de plantation uniques (avec le nombre).
- **num_point** : Polygones avec moins de 12 points (avec le nombre).
- **centroid** : Polygones dont les centroïdes sont à l'extérieur de leurs limites (avec le nombre).
- **ovlp15** : Polygones de `geo_unique` qui se chevauchent à plus de 15 %, incluant :
  - Code de plantation.
  - Superficie en hectares (3 décimales).
  - Pourcentage de chevauchement global (2 décimales).
  - Pourcentage de chevauchement de chaque polygone par rapport à l'autre (2 décimales) (avec le nombre).
- **valid** : Données filtrées à partir de `geo_unique`, avec les conditions suivantes :
  - Exclut les attributs dupliqués.
  - Supprime les polygones chevauchants (>15 %) en fonction de la superficie (si deux polygones ont la même superficie, l'un est supprimé).
  - Garantit qu'il n'y a pas de doublons dans les géométries ou les attributs après tous les filtres.
  - Inclut :
    - `area_ha` (3 décimales, 1 ha = 10 000 m²).
    - `longitude` (8 décimales).
    - `latitude` (8 décimales) (avec le nombre).
"""

# Display collapsible instructions
with st.expander("Show Instructions / Afficher les instructions", expanded=False):
    st.markdown('<div class="expander">' + (instructions_en if language == "English" else instructions_fr) + '</div>', unsafe_allow_html=True)

# File uploader for KML, Excel, or GeoJSON
file_uploader_label = "Upload a file (KML, Excel, or GeoJSON)" if language == "English" else "Téléchargez un fichier (KML, Excel ou GeoJSON)"
uploaded_file = st.file_uploader(file_uploader_label, type=["kml", "xlsx", "geojson"])

if uploaded_file is not None:
    # Progress bar initialization
    progress_text_label = "Starting analysis..." if language == "English" else "Début de l'analyse..."
    progress_text = st.empty()
    progress_bar = st.progress(0)
    progress_text.text(progress_text_label)

    # Load the data based on file type
    progress_bar.progress(10)
    progress_text.text("Loading file..." if language == "English" else "Chargement du fichier...")
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
            df['geometry'] = df['wkt_geom'].apply(loads)
            gdf = gpd.GeoDataFrame(df, geometry='geometry')
            # Set default CRS to EPSG:4326 if not defined (common for Excel files)
            if gdf.crs is None:
                gdf.set_crs(epsg=4326, inplace=True)
        elif uploaded_file.name.endswith('.geojson'):
            gdf = gpd.read_file(uploaded_file)
            # Ensure CRS is set to EPSG:4326 if not defined or different
            if gdf.crs is None or gdf.crs.to_string() != 'EPSG:4326':
                gdf = gdf.to_crs(epsg=4326)
        elif uploaded_file.name.endswith('.kml'):
            gdf = gpd.read_file(uploaded_file, driver='KML')
            # Ensure CRS is set to EPSG:4326 if not defined (KML often assumes WGS84)
            if gdf.crs is None or gdf.crs.to_string() != 'EPSG:4326':
                gdf = gdf.to_crs(epsg=4326)
        else:
            st.error("Unsupported file format!" if language == "English" else "Format de fichier non pris en charge !")
            st.stop()
    except Exception as e:
        st.error(f"Error loading file: {e}" if language == "English" else f"Erreur lors du chargement du fichier : {e}")
        st.stop()

    # Ensure required columns exist
    required_columns = ['wkt_geom', 'Plantation Code']
    if not all(col in gdf.columns for col in required_columns):
        st.error("File must contain 'wkt_geom' and 'Plantation Code' columns!" if language == "English" else "Le fichier doit contenir les colonnes 'wkt_geom' et 'Plantation Code' !")
        st.stop()

    # Reproject temporarily to EPSG:3857 for area calculation
    gdf_projected = gdf.to_crs(epsg=3857)

    # Sheet 1: All Data
    progress_bar.progress(20)
    progress_text.text("Processing all_data..." if language == "English" else "Traitement de all_data...")
    all_data = gdf.drop(columns=['geometry']).copy()
    all_data_count = len(all_data)

    # Sheet 2: Unique Geometries (geo_unique)
    progress_bar.progress(30)
    progress_text.text("Processing geo_unique..." if language == "English" else "Traitement de geo_unique...")
    geo_unique_gdf = gdf.drop_duplicates(subset=['wkt_geom']).copy()

    # Calculate num_points for geo_unique_gdf
    def count_vertices(geom):
        if isinstance(geom, Polygon):
            return len(geom.exterior.coords) - 1
        elif isinstance(geom, MultiPolygon):
            return sum(len(poly.exterior.coords) - 1 for poly in geom.geoms)
        return 0
    geo_unique_gdf['num_points'] = geo_unique_gdf['geometry'].apply(count_vertices)

    geo_unique = geo_unique_gdf.drop(columns=['geometry'])
    geo_unique_count = len(geo_unique)

    # Sheet 3: Duplicate Geometries
    progress_bar.progress(40)
    progress_text.text("Processing geo_duplicates..." if language == "English" else "Traitement de geo_duplicates...")
    geo_duplicates = gdf[gdf.duplicated(subset=['wkt_geom'], keep=False)].drop(columns=['geometry'])
    geo_duplicates_count = len(geo_duplicates)

    # Sheet 4: Duplicate Attributes (Plantation Code)
    progress_bar.progress(50)
    progress_text.text("Processing attr_duplicates..." if language == "English" else "Traitement de attr_duplicates...")
    attr_duplicates = gdf[gdf.duplicated(subset=['Plantation Code'], keep=False)][['Plantation Code']].drop_duplicates()
    attr_duplicates_count = len(attr_duplicates)

    # Sheet 5: Unique Attributes (Plantation Code)
    progress_bar.progress(60)
    progress_text.text("Processing attr_unique..." if language == "English" else "Traitement de attr_unique...")
    attr_unique = pd.DataFrame(gdf['Plantation Code'].unique(), columns=['Plantation Code'])
    attr_unique_count = len(attr_unique)

    # Sheet 6: Number of Points (vertices < 12)
    progress_bar.progress(70)
    progress_text.text("Processing num_point..." if language == "English" else "Traitement de num_point...")
    gdf['num_points'] = gdf['geometry'].apply(count_vertices)
    num_point = gdf[gdf['num_points'] < 12][['Plantation Code', 'num_points']].copy()  # Retain num_points column
    num_point_count = len(num_point)

    # Sheet 7: Centroids Outside Polygons
    progress_bar.progress(80)
    progress_text.text("Processing centroid..." if language == "English" else "Traitement de centroid...")
    gdf['centroid'] = gdf['geometry'].centroid
    gdf['centroid_outside'] = ~gdf.apply(lambda row: row['geometry'].contains(row['centroid']), axis=1)
    centroid = gdf[gdf['centroid_outside']].drop(columns=['geometry', 'centroid', 'num_points', 'centroid_outside'])
    centroid_count = len(centroid)

    # Sheet 8: Overlapping Polygons (>15% intersection) - Apply on geo_unique
    progress_bar.progress(85)
    progress_text.text("Processing ovlp15..." if language == "English" else "Traitement de ovlp15...")
    # Reproject geo_unique to EPSG:3857 for area calculation
    geo_unique_projected = geo_unique_gdf.to_crs(epsg=3857)
    overlaps = []
    for i, row1 in geo_unique_gdf.iterrows():
        for j, row2 in geo_unique_gdf.iterrows():
            if i >= j:
                continue
            geom1, geom2 = row1['geometry'], row2['geometry']
            if geom1.intersects(geom2):
                # Use projected geometries for area calculation
                geom1_projected = geo_unique_projected.loc[i, 'geometry']
                geom2_projected = geo_unique_projected.loc[j, 'geometry']
                intersection_area = geom1_projected.intersection(geom2_projected).area
                area1 = geom1_projected.area
                area2 = geom2_projected.area
                ovlp_percentage = (intersection_area / min(area1, area2)) * 100  # Overall overlap percentage
                if ovlp_percentage > 15:
                    ovlp_pct1 = (intersection_area / area1) * 100  # Overlap % of polygon 1 w.r.t. polygon 2
                    ovlp_pct2 = (intersection_area / area2) * 100  # Overlap % of polygon 2 w.r.t. polygon 1
                    overlaps.append({
                        'Plantation Code_1': row1['Plantation Code'],
                        'area_ha_1': round(area1 / 10000, 3),  # Convert to hectares and round to 3 decimals
                        'Plantation Code_2': row2['Plantation Code'],
                        'area_ha_2': round(area2 / 10000, 3),  # Convert to hectares and round to 3 decimals
                        'ovlp_percentage': round(ovlp_percentage, 2),
                        'ovlp_pct1': round(ovlp_pct1, 2),
                        'ovlp_pct2': round(ovlp_pct2, 2)
                    })
    ovlp15 = pd.DataFrame(overlaps)
    ovlp15_count = len(ovlp15)

    # Sheet 9: Valid Data - Apply on geo_unique
    progress_bar.progress(90)
    progress_text.text("Processing valid..." if language == "English" else "Traitement de valid...")
    # Start with geo_unique
    valid_gdf = geo_unique_gdf.copy()

    # Step 1: Remove duplicate attributes (Plantation Code)
    valid_gdf = valid_gdf.drop_duplicates(subset=['Plantation Code'], keep='first')

    # Step 2: Compute centroids for longitude and latitude calculation (but do not filter)
    valid_gdf['centroid'] = valid_gdf['geometry'].centroid

    # Step 3: Handle overlapping polygons (>15%) by removing based on area
    if not ovlp15.empty:
        # Rebuild plantation_code_to_index based on current valid_gdf
        plantation_code_to_index = {row['Plantation Code']: idx for idx, row in valid_gdf.iterrows()}
        indices_to_remove = set()
        for _, row in ovlp15.iterrows():
            plantation_code1 = row['Plantation Code_1']
            plantation_code2 = row['Plantation Code_2']
            area_ha_1 = row['area_ha_1']
            area_ha_2 = row['area_ha_2']
            idx1 = plantation_code_to_index.get(plantation_code1)
            idx2 = plantation_code_to_index.get(plantation_code2)
            if idx1 is not None and idx2 is not None and idx1 in valid_gdf.index and idx2 in valid_gdf.index:
                # Compare areas from ovlp15
                if area_ha_1 < area_ha_2:
                    indices_to_remove.add(idx1)
                elif area_ha_2 < area_ha_1:
                    indices_to_remove.add(idx2)
                else:  # area_ha_1 == area_ha_2
                    indices_to_remove.add(idx2)  # Remove one arbitrarily (second polygon)
        valid_gdf = valid_gdf.drop(index=indices_to_remove)

    # Step 4: Re-check for duplicates after overlap removal
    valid_gdf = valid_gdf.drop_duplicates(subset=['wkt_geom'], keep='first')
    valid_gdf = valid_gdf.drop_duplicates(subset=['Plantation Code'], keep='first')

    # Step 5: Calculate area_ha (convert area to hectares, 1 ha = 10,000 m²), longitude, and latitude
    # Reproject to EPSG:3857 for area calculation
    valid_gdf_projected = valid_gdf.to_crs(epsg=3857)
    valid_gdf['area_ha'] = valid_gdf_projected['geometry'].area / 10000  # Convert to hectares (1 ha = 10,000 m²)
    valid_gdf['area_ha'] = valid_gdf['area_ha'].round(3)  # Round to 3 decimal places
    valid_gdf['longitude'] = valid_gdf['centroid'].x.round(8)  # Longitude in WGS84 degrees
    valid_gdf['latitude'] = valid_gdf['centroid'].y.round(8)  # Latitude in WGS84 degrees

    # Drop unnecessary columns for the valid sheet, handling missing columns
    columns_to_drop = ['geometry', 'centroid']
    valid = valid_gdf.drop(columns=[col for col in columns_to_drop if col in valid_gdf.columns])
    valid_count = len(valid)

    # Generate Excel report
    progress_bar.progress(95)
    progress_text.text("Generating Excel report..." if language == "English" else "Génération du rapport Excel...")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        all_data.to_excel(writer, sheet_name=f'all_data ({all_data_count})', index=False)
        geo_unique.to_excel(writer, sheet_name=f'geo_unique ({geo_unique_count})', index=False)
        geo_duplicates.to_excel(writer, sheet_name=f'geo_duplicates ({geo_duplicates_count})', index=False)
        attr_duplicates.to_excel(writer, sheet_name=f'attr_duplicates ({attr_duplicates_count})', index=False)
        attr_unique.to_excel(writer, sheet_name=f'attr_unique ({attr_unique_count})', index=False)
        num_point.to_excel(writer, sheet_name=f'num_point ({num_point_count})', index=False)
        centroid.to_excel(writer, sheet_name=f'centroid ({centroid_count})', index=False)
        ovlp15.to_excel(writer, sheet_name=f'ovlp15 ({ovlp15_count})', index=False)
        valid.to_excel(writer, sheet_name=f'valid ({valid_count})', index=False)

    # Complete progress
    progress_bar.progress(100)
    progress_text.text("Analysis complete!" if language == "English" else "Analyse terminée !")

    # Prepare the download file name
    uploaded_file_name = uploaded_file.name.split('.')[0]  # Remove the extension
    uploaded_file_name = uploaded_file_name.replace(' ', '_')  # Replace spaces with underscores
    report_file_name = f"geospatial_report_{uploaded_file_name}.xlsx"

    # Provide download button for the Excel report
    download_label = "Download Excel Report" if language == "English" else "Télécharger le rapport Excel"
    st.download_button(
        label=download_label,
        data=output.getvalue(),
        file_name=report_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Processing complete! Download the report using the button above." if language == "English" else "Traitement terminé ! Téléchargez le rapport en utilisant le bouton ci-dessus.")
else:
    st.info("Please upload a file to begin." if language == "English" else "Veuillez télécharger un fichier pour commencer.")
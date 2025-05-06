import streamlit as st
import geopandas as gpd
import pandas as pd
from shapely.wkt import loads
from shapely.geometry import Polygon, MultiPolygon
import io
from openpyxl import Workbook

# Streamlit app title (in both languages)
st.title("Geospatial Data Preprocessing App")

# Language selection
language = st.selectbox("Choose language / Choisir la langue", ["English", "French"])

# Instructions in English and French
instructions_en = """
### Instructions
- **File Requirements**: Upload a file in KML, Excel (.xlsx), or GeoJSON format. The file must contain a column named `wkt_geom` with Polygon or MultiPolygon geometries in WKT format, and a column named `Plantation Code` for attribute analysis.
- **Output Tabs in the Excel Report**:
  - **all_data**: All loaded data (with count in parentheses).
  - **geo_unique**: Unique geometries (with count).
  - **geo_duplicates**: Duplicate geometries (with count).
  - **attr_duplicates**: Duplicate Plantation Codes (with count).
  - **attr_unique**: Unique Plantation Codes (with count).
  - **num_point**: Polygons with fewer than 12 points (with count).
  - **centroid**: Polygons whose centroids are outside their boundaries (with count).
  - **ovlp15**: Polygons overlapping by more than 15% (with count).
"""

instructions_fr = """
### Instructions
- **Exigences du fichier** : Téléchargez un fichier au format KML, Excel (.xlsx) ou GeoJSON. Le fichier doit contenir une colonne nommée `wkt_geom` avec des géométries de type Polygon ou MultiPolygon au format WKT, ainsi qu'une colonne nommée `Plantation Code` pour l'analyse des attributs.
- **Onglets de sortie dans le rapport Excel** :
  - **all_data** : Toutes les données chargées (avec le nombre entre parenthèses).
  - **geo_unique** : Géométries uniques (avec le nombre).
  - **geo_duplicates** : Géométries dupliquées (avec le nombre).
  - **attr_duplicates** : Codes de plantation dupliqués (avec le nombre).
  - **attr_unique** : Codes de plantation uniques (avec le nombre).
  - **num_point** : Polygones avec moins de 12 points (avec le nombre).
  - **centroid** : Polygones dont les centroïdes sont à l'extérieur de leurs limites (avec le nombre).
  - **ovlp15** : Polygones qui se chevauchent à plus de 15 % (avec le nombre).
"""

# Display instructions based on selected language
if language == "English":
    st.markdown(instructions_en)
else:
    st.markdown(instructions_fr)

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
        elif uploaded_file.name.endswith('.geojson'):
            gdf = gpd.read_file(uploaded_file)
        elif uploaded_file.name.endswith('.kml'):
            gdf = gpd.read_file(uploaded_file, driver='KML')
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

    # Sheet 1: All Data
    progress_bar.progress(20)
    progress_text.text("Processing all_data..." if language == "English" else "Traitement de all_data...")
    all_data = gdf.drop(columns=['geometry']).copy()
    all_data_count = len(all_data)

    # Sheet 2: Unique Geometries
    progress_bar.progress(30)
    progress_text.text("Processing geo_unique..." if language == "English" else "Traitement de geo_unique...")
    geo_unique = gdf.drop_duplicates(subset=['wkt_geom']).drop(columns=['geometry'])
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
    def count_vertices(geom):
        if isinstance(geom, Polygon):
            return len(geom.exterior.coords) - 1
        elif isinstance(geom, MultiPolygon):
            return sum(len(poly.exterior.coords) - 1 for poly in geom.geoms)
        return 0

    gdf['num_points'] = gdf['geometry'].apply(count_vertices)
    num_point = gdf[gdf['num_points'] < 12].drop(columns=['geometry'])
    num_point_count = len(num_point)

    # Sheet 7: Centroids Outside Polygons
    progress_bar.progress(80)
    progress_text.text("Processing centroid..." if language == "English" else "Traitement de centroid...")
    gdf['centroid'] = gdf['geometry'].centroid
    gdf['centroid_outside'] = ~gdf.apply(lambda row: row['geometry'].contains(row['centroid']), axis=1)
    centroid = gdf[gdf['centroid_outside']].drop(columns=['geometry', 'centroid', 'num_points', 'centroid_outside'])
    centroid_count = len(centroid)

    # Sheet 8: Overlapping Polygons (>15% intersection)
    progress_bar.progress(90)
    progress_text.text("Processing ovlp15..." if language == "English" else "Traitement de ovlp15...")
    overlaps = []
    for i, row1 in gdf.iterrows():
        for j, row2 in gdf.iterrows():
            if i >= j:
                continue
            geom1, geom2 = row1['geometry'], row2['geometry']
            if geom1.intersects(geom2):
                intersection_area = geom1.intersection(geom2).area
                area1 = geom1.area
                overlap_percentage = (intersection_area / area1) * 100
                if overlap_percentage > 15:
                    overlaps.append({
                        'ID_1': i,
                        'ID_2': j,
                        'Overlap_Percentage': overlap_percentage,
                        'Intersection_Area': intersection_area
                    })
    ovlp15 = pd.DataFrame(overlaps)
    ovlp15_count = len(ovlp15)

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

    # Complete progress
    progress_bar.progress(100)
    progress_text.text("Analysis complete!" if language == "English" else "Analyse terminée !")

    # Provide download button for the Excel report
    download_label = "Download Excel Report" if language == "English" else "Télécharger le rapport Excel"
    st.download_button(
        label=download_label,
        data=output.getvalue(),
        file_name="geospatial_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Processing complete! Download the report using the button above." if language == "English" else "Traitement terminé ! Téléchargez le rapport en utilisant le bouton ci-dessus.")
else:
    st.info("Please upload a file to begin." if language == "English" else "Veuillez télécharger un fichier pour commencer.")
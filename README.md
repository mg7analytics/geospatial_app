# Geospatial Data Preprocessing Application

This is a Streamlit-based application designed to preprocess geospatial data from KML, Excel (.xlsx), or GeoJSON files and generate a detailed Excel report. The app handles Polygon and MultiPolygon geometries, analyzes unique and duplicate data, and identifies specific geometric properties such as point counts and overlaps.

## Features
- Supports KML, Excel (.xlsx), and GeoJSON file formats.
- Processes `wkt_geom` column with Polygon or MultiPolygon geometries in WKT format.
- Generates an Excel report with the following sheets:
  - `all_data (count)`: All loaded data.
  - `geo_unique (count)`: Unique geometries.
  - `geo_duplicates (count)`: Duplicate geometries.
  - `attr_duplicates (count)`: Duplicate Plantation Codes.
  - `attr_unique (count)`: Unique Plantation Codes.
  - `num_point (count)`: Geometries with fewer than 12 points.
  - `centroid (count)`: Geometries whose centroids are outside their boundaries.
  - `ovlp15 (count)`: Geometries overlapping by more than 15%.
- Includes a progress bar during analysis.
- Offers language selection (English/French) for instructions and UI elements.
- Provides a downloadable Excel report.

## Requirements
- Python 3.7 or higher
- Dependencies listed in `requirements.txt`:
  - `streamlit`
  - `geopandas`
  - `shapely`
  - `pandas`
  - `openpyxl`

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/mg7analytics/geospatial_app.git
   cd geospatial_app
   ```
2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Run the Streamlit app:
   ```bash
   streamlit run geospatial_app.py
   ```
2. Open your browser and navigate to the provided local URL (e.g., `http://localhost:8501`).
3. Select your preferred language (English or French) from the dropdown menu.
4. Upload a file (KML, Excel, or GeoJSON) containing the required columns:
   - `wkt_geom`: Column with WKT-formatted Polygon or MultiPolygon geometries.
   - `Plantation Code`: Column for attribute analysis.
5. Wait for the analysis to complete (progress bar will update).
6. Download the generated Excel report using the provided button.

## File Requirements
- The input file must include a `wkt_geom` column with valid WKT strings for Polygon or MultiPolygon geometries.
- The input file must include a `Plantation Code` column for attribute analysis.
- Supported formats: KML, Excel (.xlsx), GeoJSON.

## Contributing
Feel free to fork this repository, submit issues, or create pull requests to improve the application. Contributions are welcome!

## License
This project is licensed under the [MIT License](LICENSE). (Note: Add a `LICENSE` file to your repository if you choose this license or specify your preferred license.)

## Contact
For questions or support, please open an issue on the GitHub repository or contact the maintainer at mg7.analytics@gmail.com.

## Acknowledgments
- Built with [Streamlit](https://streamlit.io/) for the web interface.
- Utilizes [GeoPandas](https://geopandas.org/) and [Shapely](https://shapely.readthedocs.io/) for geospatial processing.
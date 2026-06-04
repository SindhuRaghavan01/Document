
import geopandas as gpd
import numpy as np
import pandas as pd
from shapely.geometry import Point, Polygon
import matplotlib.pyplot as plt
from math import radians, cos, sin, asin, sqrt

# equidistance points on line shape

def equidistance_points(line_shape_file, num_points):

    gdf = gpd.read_file(line_shape_file)
#     gdf.plot()
    line = gdf.geometry.iloc[0]

    total_length = line.length
    equidist = total_length / (num_points + 1)

    equidistance_points = []
    for i in range(1, num_points + 1):
        distance_along_line = i * equidist
        point = line.interpolate(distance_along_line)
        longitude, latitude = point.x, point.y
        equidistance_points.append((longitude, latitude))

    return equidistance_points


def haversine(lon1, lat1, lon2, lat2):
    """
    Calculate the great circle distance in kilometers between two points 
    on the earth (specified in decimal degrees)
    """
    # convert decimal degrees to radians 
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])

    # haversine formula 
    dlon = lon2 - lon1 
    dlat = lat2 - lat1 
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a)) 
    r = 6371 # Radius of earth in kilometers. Use 3956 for miles. Determines return value units.
    return c * r


# Maximum radius circles

def max_radius_circle(polygon_shape_file, line_shape_file, num_points):

    gdf = gpd.read_file(polygon_shape_file)
#     gdf.plot()
    polygon = gdf.geometry.iloc[0]

    long_lat_list = equidistance_points(line_shape_file, num_points)

    radius_list = []
    for long_lat in long_lat_list:
        long, lat = long_lat
        point = Point(long, lat)
        if polygon.contains(point):
            min_dis = np.inf
            for point_on_boundary in polygon.boundary.coords[:-1]:
                distance = point.distance(Point(point_on_boundary))
#                 point_on_boundary = Point(point_on_boundary)
#                 distance = haversine(point.x, point.y, point_on_boundary.x, point_on_boundary.y)
                min_dis = min(min_dis, distance)
            radius_list.append(min_dis)
        else:
            radius_list.append(0)

    circle_coords = {'centers': long_lat_list, 'radius': radius_list}

    return circle_coords


def find_events_passing_circles(us_map_shape_file, event_track_file, polygon_shape_file, line_shape_file, num_points):

    events = gpd.read_file(event_track_file)

    # find circle coordinates with maximum radius and it's centers
    circles_coords = max_radius_circle(
        polygon_shape_file=polygon_shape_file, line_shape_file=line_shape_file, num_points=num_points)

    centers = circles_coords['centers']
    radius = circles_coords['radius']

    passing_events = events.copy(deep=True)
    
    us_map_df = pd.read_csv(us_map_shape_file, sep='\t')
    us_map_df.dropna(axis=0, inplace=True)

    # Convert Pandas DataFrame to GeoPandas DataFrame to plot
    us_map_df['geometry'] = gpd.GeoSeries.from_wkt(us_map_df['geometry'])
    us_map_df = gpd.GeoDataFrame(us_map_df, geometry='geometry')
    
    plt.style.use('fivethirtyeight')
    fig, ax = plt.subplots(figsize=(16, 9))

    ax.set_ylim(20, 50)
    ax.set_xlim(-130, -40)
    us_map_df.plot(ax=ax, color='white', edgecolor='k')

    # set labels
#     for idx, row in us_map_df[['State', 'geometry']].iterrows():
#         ax.annotate(row['State'], xy= row['geometry'].centroid.coords[0], ha='center', fontsize=5)

    for idx, i in enumerate(centers):
        circle_polygon = Point(i).buffer(radius[idx])

        gdf_circle = gpd.GeoDataFrame(
            geometry=[circle_polygon], crs='EPSG:4326')

        gdf_circle_proj = gdf_circle.to_crs('EPSG:4326')

        circle_polygon_projected = gdf_circle_proj.geometry.iloc[0]

        x, y = circle_polygon_projected.exterior.coords.xy
        ax.plot(x, y, color='green', label=f"circle_{idx}", linewidth=0.9)
        passing_events = passing_events[passing_events.geometry.intersects(
            circle_polygon_projected)]

    passing_events.plot(ax=ax, color='red', linewidth=0.5, alpha=0.7)

    ax.set_title(f'Events passing through circles')

    fig.tight_layout()
    plt.show()
    
    passing_events['EventID'] = (passing_events['EventID']+ 27e7).astype('int64')
    
    return passing_events['EventID']
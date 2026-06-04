import cudf
import cuspatial
import geopandas
import numpy as np


def get_poly_point_intersection(polygon_gdf: geopandas.GeoDataFrame,
                                point_gdf: geopandas.GeoDataFrame,
                                batch_grid_size: int,
                                batch_poly_size: int,
                                scale: int = 5,
                                max_depth: int = 7,
                                max_size: int = 125):
    """get point-polygon intersections. The entire geodataframes can be passed to the function.
    batch sizes for both point and polygon data can be given so that processing happens in batches.

    Parameters
    ----------
    polygon_gdf : geopandas.GeoDataFrame
        _description_
    point_gdf : geopandas.GeoDataFrame
        _description_
    batch_grid_size : int
        _description_
    batch_poly_size : int
        _description_
    scale : int, optional
        _description_, by default 5
    max_depth : int, optional
        _description_, by default 7
    max_size : int, optional
        _description_, by default 125

    Returns
    -------
    _type_
        _description_
    """
    ps = []
    n_batch_grid = int(np.ceil(point_gdf.shape[0] / batch_grid_size))
    n_batch_poly = int(np.ceil(polygon_gdf.shape[0] / batch_poly_size))

    grid_sp = cuspatial.from_geopandas(point_gdf)
    polygon_sp = cuspatial.from_geopandas(polygon_gdf)
    count = 0
    xmin = [point_gdf[i * batch_grid_size: (i + 1) * batch_grid_size].geometry.x.min() for i in range(n_batch_grid)]
    xmax = [point_gdf[i * batch_grid_size: (i + 1) * batch_grid_size].geometry.x.max() for i in range(n_batch_grid)]
    ymin = [point_gdf[i * batch_grid_size: (i + 1) * batch_grid_size].geometry.y.min() for i in range(n_batch_grid)]
    ymax = [point_gdf[i * batch_grid_size: (i + 1) * batch_grid_size].geometry.y.max() for i in range(n_batch_grid)]
    # xmin, ymin, xmax, ymax

    for i in range(n_batch_grid):
        point_indices, quadtree = cuspatial.quadtree_on_points(grid_sp["geometry"][i * batch_grid_size: (i + 1) * batch_grid_size],
                                                               xmin[i],
                                                               xmax[i],
                                                               ymin[i],
                                                               ymax[i],
                                                               scale,
                                                               max_depth,
                                                               max_size)

        for j in range(n_batch_poly):
            polygons = polygon_sp["geometry"][j * batch_poly_size: (j + 1) * batch_poly_size]

            poly_bboxes = cuspatial.polygon_bounding_boxes(
                polygons
            )
            intersections = cuspatial.join_quadtree_and_bounding_boxes(
                quadtree,
                poly_bboxes,
                polygons.polygons.x.min(),
                polygons.polygons.x.max(),
                polygons.polygons.y.min(),
                polygons.polygons.y.max(),
                scale,
                max_depth
            )
            polygons_and_points = cuspatial.quadtree_point_in_polygon(
                intersections,
                quadtree,
                point_indices,
                grid_sp["geometry"][i * batch_grid_size: (i + 1) * batch_grid_size],
                polygons
            )
            polygons_and_points["point_index"] = polygons_and_points["point_index"] + i * batch_grid_size
            polygons_and_points["polygon_index"] = polygons_and_points["polygon_index"] + j * batch_poly_size
            if count == 0:
                ps = polygons_and_points
            else:
                ps = cudf.concat([ps, polygons_and_points])
            count += 1
    return ps

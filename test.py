from qgis.core import *
from qgis.analysis import QgsZonalStatistics

# supply path to qgis install location
QgsApplication.setPrefixPath("C:/PROGRA~1/QGIS3~1.8/apps/qgis", True)
# create a reference to the QgsApplication, setting the second argument to False disables the GUI
qgs = QgsApplication([], False)
# load providers
qgs.initQgis()


# Write your code here to load some layers, use processing , algorithms, etc.
rasterPath = r"C:\Users\codhy\Desktop\H8\data\MVC_test\H8_NDVI_20190601_reclass_clip.tif"
vLayer = QgsVectorLayer(r"C:\Users\codhy\Desktop\H8\data\MVC_test\china_province.shp", "my_vector", "ogr")
rLayer = QgsRasterLayer(rasterPath, "my_raster")
if not (vLayer.isValid() and rLayer.isValid()):
    print("Error loading layers...")
band = 1
zonalstats = QgsZonalStatistics(vLayer, rLayer,  'ipv_', band, QgsZonalStatistics.Mean)
zonalstats.calculateStatistics(None)



# for field in vLayer.fields():
#     print(field.name(), field.typeName())
# # for feature in vLayer.getFeatures():
# #     print(feature[0], feature[1], feature[9])
#
# caps = vLayer.dataProvider().capabilities()
# # Check if a particular capability is supported:
# # if caps & QgsVectorDataProvider.DeleteFeatures:
# #     print('The layer supports DeleteFeatures')
# #     res = vLayer.dataProvider().deleteFeatures([5, 10])
#
# fid = 10 # ID of the feature we will modify
# if caps & QgsVectorDataProvider.ChangeAttributeValues:
#     attrs = { 3 : 2}
#     vLayer.dataProvider().changeAttributeValues({ fid : attrs })
# When your script is complete, call exitQgis() to remove the provider and layer registries from memory
qgs.exitQgis()

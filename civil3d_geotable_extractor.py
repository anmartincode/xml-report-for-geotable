"""
Civil 3D Rail Alignment Geotable Data Extractor for Dynamo
This script extracts rail alignment geotable data from Civil 3D documents
and prepares it for XML report generation.

Usage in Dynamo:
- Import this script as a Python Script node
- Connect Civil 3D document and alignment inputs
- Outputs structured data ready for XML formatting
"""

import clr
import sys

# Add Civil 3D API references - only essential ones for Dynamo
try:
    clr.AddReference('AeccDbMgd')
    clr.AddReference('AcMgd')
    clr.AddReference('AcCoreMgd')
    clr.AddReference('AcDbMgd')
except:
    pass  # Some references may already be loaded

# Import Civil 3D namespaces
from Autodesk.Civil.ApplicationServices import *
from Autodesk.Civil.DatabaseServices import *
from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.DatabaseServices import *
from Autodesk.AutoCAD.Geometry import *
from System.Collections.Generic import List
from System import *

# Import Reference for out parameters
try:
    from clr import Reference
except:
    Reference = None  # Fallback if not available


class GeoTableDataExtractor:
    """Extract geotable data from Civil 3D rail alignments"""
    
    def __init__(self, document=None):
        """Initialize the extractor with Civil 3D document"""
        if document is None:
            self.civil_doc = CivilApplication.ActiveDocument
        else:
            self.civil_doc = document
        
        self.acad_doc = Application.DocumentManager.MdiActiveDocument
        self.db = self.acad_doc.Database
        
    def get_all_alignments(self):
        """Get all alignments from the Civil 3D document"""
        alignments = []
        
        try:
            trans = self.db.TransactionManager.StartTransaction()
            
            # Get alignment collection
            alignment_ids = self.civil_doc.GetAlignmentIds()
            
            for oid in alignment_ids:
                alignment = trans.GetObject(oid, OpenMode.ForRead)
                
                # Safely get properties with fallbacks and convert to Python types
                try:
                    start_station = float(alignment.StartingStation)
                except:
                    start_station = 0.0
                
                try:
                    end_station = float(alignment.EndingStation)
                except:
                    try:
                        end_station = float(alignment.Length) if hasattr(alignment, 'Length') else 0.0
                    except:
                        end_station = 0.0
                
                # Convert all .NET types to Python types
                alignments.append({
                    'id': str(oid),
                    'name': str(alignment.Name) if hasattr(alignment, 'Name') else 'Unknown',
                    'description': str(alignment.Description) if hasattr(alignment, 'Description') else '',
                    'length': float(alignment.Length) if hasattr(alignment, 'Length') else 0.0,
                    'start_station': start_station,
                    'end_station': end_station,
                    'object': alignment
                })
            
            trans.Commit()
            trans.Dispose()
            
        except Exception as e:
            if trans:
                trans.Abort()
                trans.Dispose()
            raise Exception("Error retrieving alignments: " + str(e))
        
        return alignments
    
    def extract_alignment_stations(self, alignment, interval=None):
        """Extract station data along an alignment"""
        stations_data = []
        
        if interval is None:
            interval = 10.0  # Default 10 unit interval
        
        try:
            start_station = float(str(alignment.StartingStation))
            end_station = float(str(alignment.EndingStation))
            current_station = start_station
            
            # Sample stations along the alignment
            while current_station <= end_station:
                try:
                    # Create a simple station entry with Python native types
                    station_data = {
                        'station': float(str(current_station)),
                        'x': 0.0,  # Will try to get actual coordinates
                        'y': 0.0,
                        'z': 0.0,
                        'direction': 0.0,
                        'offset': 0.0
                    }
                    
                    # Try to get actual point coordinates using different methods
                    try:
                        # Method 1: Try GetStationOffsetElevationAtXY (inverse might exist)
                        # Actually, let's use the polyline representation
                        polyline = alignment.GetPolyline()
                        if polyline:
                            # Get a point along the polyline
                            # This is approximate but will work
                            param = float(str((current_station - start_station) / (end_station - start_station)))
                            point = polyline.GetPointAtParameter(param * float(str(polyline.EndParam)))
                            station_data['x'] = float(str(point.X))
                            station_data['y'] = float(str(point.Y))
                            station_data['z'] = float(str(point.Z))
                    except:
                        # If we can't get coordinates, at least record the station
                        pass
                    
                    stations_data.append(station_data)
                    
                except Exception as e:
                    # Skip problematic stations
                    pass
                
                current_station += interval
            
            # Always include end station
            if len(stations_data) > 0 and abs(stations_data[-1]['station'] - float(str(end_station))) > 0.001:
                try:
                    station_data = {
                        'station': float(str(end_station)),
                        'x': 0.0,
                        'y': 0.0,
                        'z': 0.0,
                        'direction': 0.0,
                        'offset': 0.0
                    }
                    
                    try:
                        polyline = alignment.GetPolyline()
                        if polyline:
                            point = polyline.EndPoint
                            station_data['x'] = float(str(point.X))
                            station_data['y'] = float(str(point.Y))
                            station_data['z'] = float(str(point.Z))
                    except:
                        pass
                    
                    stations_data.append(station_data)
                except:
                    pass
                    
        except Exception as e:
            # Return what we have even if there's an error
            pass
        
        return stations_data
    
    def extract_alignment_curves(self, alignment):
        """Extract curve data from alignment entities"""
        curves_data = []
        
        try:
            # Get alignment entities (tangents, curves, spirals)
            entities = alignment.Entities
            
            # Use Python-style iteration instead of .NET Item indexing
            entity_index = 0
            for entity in entities:
                try:
                    # Safely get entity type
                    try:
                        entity_type = str(entity.EntityType)
                    except:
                        entity_type = 'Unknown'
                    
                    # Build entity data with very safe property access
                    entity_data = {
                        'index': entity_index,
                        'type': entity_type
                    }
                    
                    # Try to get basic properties
                    try:
                        entity_data['start_station'] = float(entity.StartStation)
                    except:
                        entity_data['start_station'] = 0.0
                    
                    try:
                        entity_data['end_station'] = float(entity.EndStation)
                    except:
                        entity_data['end_station'] = 0.0
                    
                    try:
                        entity_data['length'] = float(entity.Length)
                    except:
                        entity_data['length'] = 0.0
                    
                    # Add type-specific data with safe property access
                    if 'Arc' in entity_type or 'Curve' in entity_type:
                        try:
                            entity_data['radius'] = float(entity.Radius)
                        except:
                            entity_data['radius'] = 0.0
                        try:
                            entity_data['chord_length'] = float(entity.ChordLength)
                        except:
                            entity_data['chord_length'] = 0.0
                        try:
                            entity_data['direction'] = 'CW' if entity.Clockwise else 'CCW'
                        except:
                            entity_data['direction'] = 'Unknown'
                        
                    elif 'Spiral' in entity_type:
                        try:
                            entity_data['radius_in'] = float(entity.RadiusIn)
                        except:
                            entity_data['radius_in'] = 0.0
                        try:
                            entity_data['radius_out'] = float(entity.RadiusOut)
                        except:
                            entity_data['radius_out'] = 0.0
                        try:
                            entity_data['spiral_type'] = str(entity.SpiralDefinition)
                        except:
                            entity_data['spiral_type'] = 'Unknown'
                        
                    elif 'Tangent' in entity_type or 'Line' in entity_type:
                        try:
                            entity_data['bearing'] = float(entity.Direction)
                        except:
                            entity_data['bearing'] = 0.0
                    
                    curves_data.append(entity_data)
                    entity_index += 1
                    
                except Exception as e:
                    # Skip problematic entities but continue
                    entity_index += 1
                    pass
                
        except Exception as e:
            # Return what we have even if there's an error
            pass
        
        return curves_data
    
    def extract_superelevation_data(self, alignment):
        """Extract superelevation data if available"""
        superelevation_data = []
        
        try:
            if hasattr(alignment, 'SuperelevationData') and alignment.SuperelevationData is not None:
                se_data = alignment.SuperelevationData
                
                # Extract critical stations if they exist
                if hasattr(se_data, 'CriticalStations'):
                    for critical_station in se_data.CriticalStations:
                        try:
                            superelevation_data.append({
                                'station': getattr(critical_station, 'Station', 0.0),
                                'type': critical_station.TransitionType.ToString() if hasattr(critical_station, 'TransitionType') else 'Unknown',
                                'left_slope': getattr(critical_station, 'LeftOutsideSlope', 0.0),
                                'right_slope': getattr(critical_station, 'RightOutsideSlope', 0.0)
                            })
                        except:
                            # Skip problematic critical stations
                            pass
        except:
            # Superelevation might not be available - this is fine
            pass
        
        return superelevation_data
    
    def extract_geotable_data(self, alignment_name=None, station_interval=10.0):
        """
        Main method to extract complete geotable data for an alignment
        
        Parameters:
        - alignment_name: Name of specific alignment (None for all alignments)
        - station_interval: Interval for station sampling
        
        Returns:
        - Dictionary containing all geotable data
        """
        geotable_data = {
            'project_name': str(self.acad_doc.Name),
            'timestamp': str(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
            'alignments': []
        }
        
        alignments = self.get_all_alignments()
        
        # Filter by name if specified
        if alignment_name:
            alignments = [a for a in alignments if a['name'] == alignment_name]
        
        for alignment_info in alignments:
            alignment_obj = alignment_info['object']
            
            # Convert all .NET types to Python types explicitly
            # DON'T include the 'object' itself - it's a .NET object that can't be serialized
            alignment_data = {
                'name': str(alignment_info['name']),
                'description': str(alignment_info['description']),
                'id': str(alignment_info['id']),
                'length': float(alignment_info['length']),
                'start_station': float(alignment_info['start_station']),
                'end_station': float(alignment_info['end_station']),
                'stations': self.extract_alignment_stations(alignment_obj, station_interval),
                'curves': self.extract_alignment_curves(alignment_obj),
                'superelevation': self.extract_superelevation_data(alignment_obj)
            }
            
            # Note: we don't include alignment_info['object'] because it's a .NET object
            
            geotable_data['alignments'].append(alignment_data)
        
        return geotable_data


# ============= DYNAMO SCRIPT ENTRY POINT =============
# The code below this line is executed when run in Dynamo

def sanitize_for_python(obj):
    """Recursively convert all .NET types to Python native types"""
    if obj is None:
        return None
    
    # Check if already a Python dict (most important check first!)
    if isinstance(obj, dict):
        py_dict = {}
        for k, v in obj.items():
            py_dict[str(k)] = sanitize_for_python(v)
        return py_dict
    
    # Check if already a Python list
    if isinstance(obj, list):
        return [sanitize_for_python(item) for item in obj]
    
    # Check if already a Python tuple
    if isinstance(obj, tuple):
        return tuple(sanitize_for_python(item) for item in obj)
    
    # Handle booleans (before numbers!)
    if isinstance(obj, bool):
        return bool(obj)
    
    # Handle strings
    if isinstance(obj, str):
        return str(obj)
    
    # Handle numbers (int and float)
    if isinstance(obj, (int, float)) and not isinstance(obj, bool):
        try:
            return float(str(obj))
        except:
            return obj
    
    # Now handle .NET dict-like objects (have both 'keys' and 'items' methods)
    # Check BOTH to ensure it's really dict-like and not just iterable
    if hasattr(obj, 'keys') and hasattr(obj, 'items'):
        try:
            if callable(getattr(obj, 'items')) and callable(getattr(obj, 'keys')):
                py_dict = {}
                for k, v in obj.items():
                    py_dict[str(k)] = sanitize_for_python(v)
                return py_dict
        except:
            pass
    
    # Handle .NET list-like objects (iterable but not string/dict)
    # But first make sure it's NOT dict-like!
    if not hasattr(obj, 'keys'):
        try:
            # Try to iterate
            py_list = []
            for item in obj:
                py_list.append(sanitize_for_python(item))
            return py_list
        except (TypeError, AttributeError):
            pass
    
    # Last resort: convert to string
    return str(obj)


def main(alignment_name=None, station_interval=10.0):
    """
    Main function called by Dynamo
    
    Inputs (IN[0], IN[1]):
    - alignment_name: String - Name of alignment to extract (None for all)
    - station_interval: Float - Station sampling interval in drawing units
    
    Output (OUT):
    - Dictionary containing all extracted geotable data
    """
    try:
        extractor = GeoTableDataExtractor()
        result = extractor.extract_geotable_data(alignment_name, station_interval)
        
        # Recursively sanitize all data to pure Python types
        sanitized_result = sanitize_for_python(result)
        
        return sanitized_result
        
    except Exception as e:
        return "Error: " + str(e)


# Check if running in Dynamo context
if 'IN' in dir():
    # Get inputs from Dynamo
    alignment_name = IN[0] if len(IN) > 0 and IN[0] != "" else None
    station_interval = IN[1] if len(IN) > 1 else 10.0
    
    # Execute and return output
    OUT = main(alignment_name, station_interval)
else:
    # Standalone execution for testing
    OUT = "Script loaded successfully. Use in Dynamo context."


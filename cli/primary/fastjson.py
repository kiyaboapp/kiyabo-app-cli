import re
from typing import List, Dict, Any, Union, Optional

class FastJSON:
    """
    ====================================================================
    Class Module Name: FastJSON
    ONE CLASS - SIMPLE AND WORKING - DEEP NESTING FIXED
    November 15, 2025 - Arusha, Tanzania
    ====================================================================
    """
    
    class JSONItem:
        def __init__(self):
            self.key = ""
            self.value = ""
            self.ItemType = ""  # "V"=value, "A"=array, "O"=object
    
    def __init__(self):
        self.mItems = [self.JSONItem() for _ in range(10)]
        self.mCount = 0
    
    def _nz(self, value, default=""):
        """Helper function to handle None values"""
        return value if value is not None else default
    
    # ====================================================================
    # Add simple value
    # ====================================================================
    def Add(self, key: str, value: Any):
        self.AddItem(key, str(self._nz(value, "")), "V")
    
    # ====================================================================
    # Add array
    # ====================================================================
    def AddArray(self, key: str, *values):
        arrStr = ""
        
        for i in range(len(values)):
            if i > 0:
                arrStr += "|"
            arrStr += str(self._nz(values[i], ""))
        
        self.AddItem(key, arrStr, "A")
    
    # ====================================================================
    # Add nested object
    # ====================================================================
    def AddObject(self, key: str, Obj: 'FastJSON'):
        self.AddItem(key, Obj.ToRaw(), "O")
    
    # ====================================================================
    # Get value (supports dot notation)
    # ====================================================================
    def GetValue(self, path: str) -> str:
        parts = path.split('.')
        
        # Get first level
        idx = self.FindKey(parts[0])
        if idx == -1:
            return ""
        
        # Simple value at root level
        if self.mItems[idx].ItemType == "V":
            if len(parts) == 1:
                return self.mItems[idx].value
            return ""
        
        # Array at root level (can't navigate into arrays)
        if self.mItems[idx].ItemType == "A":
            return ""
        
        # Object navigation
        if self.mItems[idx].ItemType == "O":
            # If just requesting the object key itself (no dots)
            if len(parts) == 1:
                return ""
            
            # Parse the nested object
            tempObj = FastJSON()
            tempObj.Parse(self.mItems[idx].value)
            
            # Build remaining path and recurse
            remainingPath = ""
            for i in range(1, len(parts)):
                if i > 1:
                    remainingPath += "."
                remainingPath += parts[i]
            
            # Recursive call
            return tempObj.GetValue(remainingPath)
        
        return ""
    
    # ====================================================================
    # Get values (supports dot notation for nested objects)
    # Returns array of values; objects return "[Object]", arrays return "[Array]"
    # ====================================================================
    def GetValues(self, path: str = "") -> list:
        # If no path specified, return root values
        if len(path) == 0:
            if self.mCount == 0:
                return []
            
            result = []
            for i in range(self.mCount):
                if self.mItems[i].ItemType == "V":
                    result.append(self.mItems[i].value)
                elif self.mItems[i].ItemType == "A":
                    result.append("[Array]")
                elif self.mItems[i].ItemType == "O":
                    result.append("[Object]")
            
            return result
        
        # Get values from nested object
        parts = path.split('.')
        idx = self.FindKey(parts[0])
        
        if idx == -1:
            return []
        
        # If not an object type, return empty array
        if self.mItems[idx].ItemType != "O":
            return []
        
        # Parse the object
        tempObj = FastJSON()
        tempObj.Parse(self.mItems[idx].value)
        
        # If single level path, return its values
        if len(parts) == 1:
            return tempObj.GetValues()
        
        # Build remaining path for deeper nesting
        remainingPath = ""
        for i in range(1, len(parts)):
            if i > 1:
                remainingPath += "."
            remainingPath += parts[i]
        
        # Recursive call
        return tempObj.GetValues(remainingPath)
    
    # ====================================================================
    # Get array
    # ====================================================================
    def GetArray(self, path: str) -> list:
        parts = path.split('.')
        
        # Get first level
        idx = self.FindKey(parts[0])
        if idx == -1:
            return []
        
        if len(parts) == 1:
            # Direct array
            if self.mItems[idx].ItemType == "A":
                if len(self.mItems[idx].value) == 0:
                    return []
                else:
                    return self.mItems[idx].value.split('|')
            else:
                return []
        else:
            # Nested array
            if self.mItems[idx].ItemType == "O":
                tempObj = FastJSON()
                tempObj.Parse(self.mItems[idx].value)
                
                # Build remaining path
                remainingPath = ""
                for i in range(1, len(parts)):
                    if i > 1:
                        remainingPath += "."
                    remainingPath += parts[i]
                
                return tempObj.GetArray(remainingPath)
            else:
                return []
    
    # ====================================================================
    # Get nested object
    # ====================================================================
    def GetObject(self, path: str) -> Optional['FastJSON']:
        parts = path.split('.')
        
        idx = self.FindKey(parts[0])
        if idx == -1:
            return None
        
        if self.mItems[idx].ItemType != "O":
            return None
        
        tempObj = FastJSON()
        tempObj.Parse(self.mItems[idx].value)
        
        if len(parts) == 1:
            # Return this object
            return tempObj
        else:
            # Navigate deeper
            remainingPath = ""
            for i in range(1, len(parts)):
                if i > 1:
                    remainingPath += "."
                remainingPath += parts[i]
            
            return tempObj.GetObject(remainingPath)
    
    # ====================================================================
    # Check if key exists
    # ====================================================================
    def HasKey(self, path: str) -> bool:
        parts = path.split('.')
        
        idx = self.FindKey(parts[0])
        
        if idx == -1:
            return False
        
        if len(parts) == 1:
            return True
        
        # Check nested path
        if self.mItems[idx].ItemType == "O":
            tempObj = FastJSON()
            tempObj.Parse(self.mItems[idx].value)
            
            remainingPath = ""
            for i in range(1, len(parts)):
                if i > 1:
                    remainingPath += "."
                remainingPath += parts[i]
            
            return tempObj.HasKey(remainingPath)
        else:
            return False
    
    # ====================================================================
    # Get keys (supports dot notation for nested objects)
    # ====================================================================
    def GetKeys(self, path: str = "") -> list:
        # If no path specified, return root keys
        if len(path) == 0:
            if self.mCount == 0:
                return []
            
            result = []
            for i in range(self.mCount):
                result.append(self.mItems[i].key)
            
            return result
        
        # Get keys from nested object
        parts = path.split('.')
        idx = self.FindKey(parts[0])
        
        if idx == -1:
            return []
        
        # If not an object type, return empty array
        if self.mItems[idx].ItemType != "O":
            return []
        
        # Parse the object
        tempObj = FastJSON()
        tempObj.Parse(self.mItems[idx].value)
        
        # If single level path, return its keys
        if len(parts) == 1:
            return tempObj.GetKeys()
        
        # Build remaining path for deeper nesting
        remainingPath = ""
        for i in range(1, len(parts)):
            if i > 1:
                remainingPath += "."
            remainingPath += parts[i]
        
        # Recursive call
        return tempObj.GetKeys(remainingPath)
    
    # ====================================================================
    # Count items
    # ====================================================================
    def Count(self) -> int:
        return self.mCount
    
    # ====================================================================
    # Update value
    # ====================================================================
    def UpdateValue(self, key: str, NewValue: Any) -> bool:
        idx = self.FindKey(key)
        
        if idx != -1 and self.mItems[idx].ItemType == "V":
            self.mItems[idx].value = str(self._nz(NewValue, ""))
            return True
        
        return False
    
    # ====================================================================
    # Delete key
    # ====================================================================
    def DeleteKey(self, key: str) -> bool:
        idx = self.FindKey(key)
        if idx == -1:
            return False
        
        # Shift items down
        for i in range(idx, self.mCount - 1):
            self.mItems[i] = self.mItems[i + 1]
        
        self.mCount -= 1
        return True
    
    # ====================================================================
    # Clear all
    # ====================================================================
    def Clear(self):
        self.mItems = [self.JSONItem() for _ in range(10)]
        self.mCount = 0
    
    # ====================================================================
    # Pretty print
    # ====================================================================
    def ToPretty(self, Indent: int = 0) -> str:
        return self.BuildPretty(Indent)
    
    # ====================================================================
    # To JSON
    # ====================================================================
    def ToJSON(self, Indent: int = 0) -> str:
        return self.BuildJSON(Indent)
    
    # ====================================================================
    # Serialize to string
    # ====================================================================
    def ToRaw(self) -> str:
        result = ""
        
        for i in range(self.mCount):
            if i > 0:
                result += chr(2)
            
            # For objects, encode their raw data to prevent Chr(1) and Chr(2) conflicts
            if self.mItems[i].ItemType == "O":
                result += self.mItems[i].key + chr(1) + self.mItems[i].ItemType + chr(1) + self.EncodeForStorage(self.mItems[i].value)
            else:
                result += self.mItems[i].key + chr(1) + self.mItems[i].ItemType + chr(1) + self.mItems[i].value
        
        return result
    
    def EncodeForStorage(self, s: str) -> str:
        # Replace Chr(1) and Chr(2) with safe control characters
        s = s.replace(chr(1), chr(3))
        s = s.replace(chr(2), chr(4))
        return s
    
    def DecodeFromStorage(self, s: str) -> str:
        # Restore Chr(1) and Chr(2) from safe control characters
        s = s.replace(chr(3), chr(1))
        s = s.replace(chr(4), chr(2))
        return s
    
    # ====================================================================
    # Deserialize from string
    # ====================================================================
    def Parse(self, rawData: str):
        self.Clear()
        
        if len(rawData) == 0:
            return
        
        pairs = rawData.split(chr(2))
        
        for i in range(len(pairs)):
            if len(pairs[i]) == 0:
                continue
            
            parts = pairs[i].split(chr(1))
            
            if len(parts) >= 3:
                # For objects, decode their stored data
                if parts[1] == "O":
                    self.AddItem(parts[0], self.DecodeFromStorage(parts[2]), parts[1])
                else:
                    self.AddItem(parts[0], parts[2], parts[1])
    
    # ====================================================================
    # SQL-safe
    # ====================================================================
    def ToSQLSafe(self) -> str:
        return self.ToRaw().replace("'", "''")
    
    # ====================================================================
    # PRIVATE HELPERS
    # ====================================================================
    def AddItem(self, key: str, value: str, ItemType: str):
        if self.mCount >= len(self.mItems):
            self.mItems.extend([self.JSONItem() for _ in range(10)])
        
        self.mItems[self.mCount].key = key
        self.mItems[self.mCount].value = value
        self.mItems[self.mCount].ItemType = ItemType
        self.mCount += 1
    
    def FindKey(self, key: str) -> int:
        for i in range(self.mCount):
            if self.mItems[i].key == key:
                return i
        return -1
    
    def BuildPretty(self, Indent: int) -> str:
        result = ""
        sp = " " * (Indent * 2)
        
        if self.mCount == 0:
            return sp + "{}"
        
        result = sp + "{" + "\n"
        
        for i in range(self.mCount):
            result += sp + '  "' + self.mItems[i].key + '": '
            
            if self.mItems[i].ItemType == "V":
                if (self.mItems[i].value.isdigit() or 
                    self.mItems[i].value.lower() == "true" or 
                    self.mItems[i].value.lower() == "false" or 
                    self.mItems[i].value.lower() == "null"):
                    result += self.mItems[i].value
                else:
                    result += '"' + self.mItems[i].value + '"'
            
            elif self.mItems[i].ItemType == "A":
                result += "[ "
                if len(self.mItems[i].value) > 0:
                    result += '"' + self.mItems[i].value.replace("|", '", "') + '"'
                result += " ]"
            
            elif self.mItems[i].ItemType == "O":
                childObj = FastJSON()
                childObj.Parse(self.mItems[i].value)
                result += "\n" + childObj.ToPretty(Indent + 1)
            
            if i < self.mCount - 1:
                result += ","
            result += "\n"
        
        result += sp + "}"
        return result
    
    def BuildJSON(self, Indent: int) -> str:
        result = ""
        sp = " " * (Indent * 2)
        
        if self.mCount == 0:
            return sp + "{}"
        
        result = sp + "{" + "\n"
        
        for i in range(self.mCount):
            result += sp + '  "' + self.EscapeJSON(self.mItems[i].key) + '": '
            
            if self.mItems[i].ItemType == "V":
                if (self.mItems[i].value.isdigit() or 
                    self.mItems[i].value.lower() == "true" or 
                    self.mItems[i].value.lower() == "false" or 
                    self.mItems[i].value.lower() == "null"):
                    result += self.mItems[i].value
                else:
                    result += '"' + self.EscapeJSON(self.mItems[i].value) + '"'
            
            elif self.mItems[i].ItemType == "A":
                result += "["
                if len(self.mItems[i].value) > 0:
                    arrParts = self.mItems[i].value.split('|')
                    for j in range(len(arrParts)):
                        if j > 0:
                            result += ", "
                        result += '"' + self.EscapeJSON(arrParts[j]) + '"'
                result += "]"
            
            elif self.mItems[i].ItemType == "O":
                childObj = FastJSON()
                childObj.Parse(self.mItems[i].value)
                result += "\n" + childObj.ToJSON(Indent + 1)
            
            if i < self.mCount - 1:
                result += ","
            result += "\n"
        
        result += sp + "}"
        return result
    
    def EscapeJSON(self, s: str) -> str:
        s = s.replace("\\", "\\\\")
        s = s.replace('"', '\\"')
        s = s.replace('\r\n', '\\n')
        s = s.replace('\r', '\\n')
        s = s.replace('\n', '\\n')
        s = s.replace('\t', '\\t')
        return s
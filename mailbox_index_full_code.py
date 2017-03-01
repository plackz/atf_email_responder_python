import win32com.client
 
DEBUG = False
EXCLUSION_LIST = [
    'SharePoint Lists - outlook',
    'Project Canendars'
]
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
 
class Oli(object):
    def __init__(self, outlook_object):
        self._obj = outlook_object
 
    def items(self):
        array_size = self._obj.Count
        for item_index in xrange(1,array_size+1):
            yield (item_index, self._obj[item_index-1])
 
    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )
 
def search_item(folders, name):
    if DEBUG: browse(folders, recursive=False)
    for index, folder in Oli(folders).items():
        if folder.Name == name:
            if DEBUG: print(" Found %s @ %d" % (folder.Name, index))
            return index, folder
    return None, None
 
def search(path):
    components = path.split('/')
    if DEBUG: print(components)
    folder = None
    root = outlook.Folders
    for name in components:
        index, folder = search_item(root, name)
        if not index:
            return None
        root = folder.Folders
 
    return folder
 
 
def browse(folders, depth=2, recursive=True):
    if not folders:
        return
    for index, folder in Oli(folders).items():
        print(" "*depth, u"(%i) [%s] [%s]" % (index, folder.Name, folder))
        if u"%s" % folder in EXCLUSION_LIST:
            continue
        if recursive:
            browse(folder.Folders, depth + 2, recursive)
 
def process_messages(folder):
    if not folder:
        print("Folder could not be found!")
        return
    messages = folder.Items
    message = messages.GetFirst()
    while message:
        # Process a message
        print("%s;%s;%s") % (message.Categories, message.Subject, message.SentOn)
        message = messages.GetNext()
 
if __name__ == "__main__":
    #list(outlook.Folders)
    f = search('Large Files\Projects\A Project')
    if DEBUG and f: print("Folder name: ", f.Name)
    process_messages(f)

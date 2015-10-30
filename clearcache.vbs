on error resume next
dim oUIResManager
dim oCache
dim oCacheElement
dim oCacheElements

set oUIResManager = createobject("UIResource.UIResourceMgr")

if oUIResManager is nothing then
wscript.quit
end if

set oCache=oUIResManager.GetCacheInfo()

if oCache is nothing then
  set oUIResManager=nothing
  wscript.quit
end if

set oCacheElements=oCache.GetCacheElements

for each oCacheElement in oCacheElements
  oCache.DeleteCacheElement(oCacheElement.CacheElementID)
next

set oCacheElements=nothing
set oUIResManager=nothing
set oCache=nothing
wscript.quit
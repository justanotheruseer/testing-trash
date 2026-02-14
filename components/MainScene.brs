sub init()
  m.state = "search"
  m.searchText = ""
  m.currentResults = []
  m.currentSources = []
  m.selectedItem = invalid
  m.settingsFocusIndex = 0
  m.ratingOptions = ["TV-G", "TV-PG", "TV-14", "TV-MA", "PG", "PG-13", "R"]
  m.searchTypes = ["movie", "series"]
  m.searchTypeIndex = 0

  m.searchLabel = m.top.findNode("searchText")
  m.searchTypeLabel = m.top.findNode("searchType")
  m.resultsGrid = m.top.findNode("resultsGrid")
  m.resultsTitle = m.top.findNode("resultsTitle")
  m.detailsPane = m.top.findNode("detailsPane")
  m.detailPoster = m.top.findNode("detailPoster")
  m.detailTitle = m.top.findNode("detailTitle")
  m.detailMeta = m.top.findNode("detailMeta")
  m.detailDescription = m.top.findNode("detailDescription")
  m.sourcesRow = m.top.findNode("sourcesRow")
  m.settingsPane = m.top.findNode("settingsPane")
  m.addonUrlText = m.top.findNode("addonUrlText")
  m.ratingText = m.top.findNode("ratingText")
  m.apiKeyText = m.top.findNode("apiKeyText")
  m.errorDialog = m.top.findNode("errorDialog")
  m.videoPlayer = m.top.findNode("videoPlayer")
  m.serviceTask = m.top.findNode("serviceTask")

  m.serviceTask.ObserveField("results", "onTaskDone")
  m.serviceTask.ObserveField("errorText", "onTaskError")

  sec = CreateObject("roRegistrySection", "StreamFinder")
  m.settings = {
    addonUrl: firstNonEmpty(sec.Read("addonUrl"), "http://localhost:7000"),
    maxRating: firstNonEmpty(sec.Read("maxRating"), "TV-14"),
    apiKey: firstNonEmpty(sec.Read("apiKey"), "")
  }

  updateSettingsLabels()
  updateSearchText()
  updateSearchType()
  m.top.SetFocus(true)
end sub

sub onTaskDone()
  result = m.serviceTask.results
  if result = invalid then return

  if result.items <> invalid
    m.currentResults = result.items
    showResults()
  else if result.streams <> invalid
    m.currentSources = result.streams
    showSources()
  end if
end sub

sub onTaskError()
  txt = m.serviceTask.errorText
  if txt <> invalid and txt <> ""
    showError(txt)
  end if
end sub

function onKeyEvent(key as string, press as boolean) as boolean
  if not press then return false

  if m.errorDialog.visible
    m.errorDialog.visible = false
    return true
  end if

  if m.videoPlayer.visible
    if key = "OK"
      togglePlayback()
      return true
    else if key = "back"
      stopPlayback()
      return true
    end if
    return false
  end if

  if m.state = "search"
    return handleSearchKeys(key)
  else if m.state = "details"
    return handleDetailsKeys(key)
  else if m.state = "settings"
    return handleSettingsKeys(key)
  end if
  return false
end function

function handleSearchKeys(key as string) as boolean
  if key = "backspace"
    if Len(m.searchText) > 0
      m.searchText = Left(m.searchText, Len(m.searchText) - 1)
      updateSearchText()
    end if
    return true
  else if key = "*"
    m.state = "settings"
    m.settingsPane.visible = true
    return true
  else if key = "up" or key = "down"
    m.searchTypeIndex = (m.searchTypeIndex + 1) mod m.searchTypes.Count()
    updateSearchType()
    return true
  else if key = "right"
    if m.currentResults.Count() > 0
      m.resultsGrid.SetFocus(true)
      return false
    end if
    return true
  else if key = "left"
    m.top.SetFocus(true)
    return true
  else if key = "OK"
    if m.currentResults.Count() > 0 and m.resultsGrid.hasFocus()
      idx = m.resultsGrid.itemSelected
      if idx >= 0 and idx < m.currentResults.Count()
        openDetails(m.currentResults[idx])
      end if
    else
      runSearch()
    end if
    return true
  else if Len(key) = 1
    m.searchText = m.searchText + key
    updateSearchText()
    return true
  end if
  return false
end function

function handleDetailsKeys(key as string) as boolean
  if key = "back"
    m.detailsPane.visible = false
    m.state = "search"
    m.top.SetFocus(true)
    return true
  else if key = "OK"
    idx = m.sourcesRow.rowItemSelected[1]
    if idx >= 0 and idx < m.currentSources.Count()
      playSource(m.currentSources[idx].url)
    end if
    return true
  end if
  return false
end function

function handleSettingsKeys(key as string) as boolean
  if key = "back"
    saveSettings()
    m.settingsPane.visible = false
    m.state = "search"
    m.top.SetFocus(true)
    return true
  else if key = "OK"
    m.settingsFocusIndex = (m.settingsFocusIndex + 1) mod 3
    if m.settingsFocusIndex = 1
      currentIndex = indexOf(m.ratingOptions, m.settings.maxRating)
      m.settings.maxRating = m.ratingOptions[(currentIndex + 1) mod m.ratingOptions.Count()]
    end if
    updateSettingsLabels()
    return true
  else if Len(key) = 1
    if m.settingsFocusIndex = 0
      m.settings.addonUrl = m.settings.addonUrl + key
    else if m.settingsFocusIndex = 2
      m.settings.apiKey = m.settings.apiKey + key
    end if
    updateSettingsLabels()
    return true
  else if key = "backspace"
    if m.settingsFocusIndex = 0 and Len(m.settings.addonUrl) > 0
      m.settings.addonUrl = Left(m.settings.addonUrl, Len(m.settings.addonUrl)-1)
    else if m.settingsFocusIndex = 2 and Len(m.settings.apiKey) > 0
      m.settings.apiKey = Left(m.settings.apiKey, Len(m.settings.apiKey)-1)
    end if
    updateSettingsLabels()
    return true
  end if
  return false
end function

sub runSearch()
  if Trim(m.searchText) = "" then
    showError("Enter search text first.")
    return
  end if
  m.serviceTask.control = "stop"
  m.serviceTask.action = "search"
  m.serviceTask.query = m.searchText
  m.serviceTask.mediaType = m.searchTypes[m.searchTypeIndex]
  m.serviceTask.apiKey = m.settings.apiKey
  m.serviceTask.maxRating = m.settings.maxRating
  m.serviceTask.control = "run"
end sub

sub openDetails(item as object)
  m.state = "details"
  m.selectedItem = item
  m.detailsPane.visible = true
  m.detailPoster.uri = item.poster
  m.detailTitle.text = item.title
  m.detailMeta.text = item.year + " • " + item.type + " • " + item.rated
  m.detailDescription.text = firstNonEmpty(item.plot, "No description available.")

  m.serviceTask.control = "stop"
  m.serviceTask.action = "sources"
  m.serviceTask.mediaType = item.type
  m.serviceTask.imdbId = item.imdbID
  m.serviceTask.addonUrl = m.settings.addonUrl
  m.serviceTask.control = "run"
end sub

sub showResults()
  content = CreateObject("roSGNode", "ContentNode")
  for each item in m.currentResults
    child = content.CreateChild("ContentNode")
    child.title = item.title
    child.poster = item.poster
  end for
  m.resultsGrid.content = content
  m.resultsTitle.visible = true
  m.resultsGrid.visible = true
  m.top.SetFocus(true)
end sub

sub showSources()
  content = CreateObject("roSGNode", "ContentNode")
  row = content.CreateChild("ContentNode")
  for each item in m.currentSources
    child = row.CreateChild("ContentNode")
    child.title = item.title
    child.subtitle = item.subtitle
  end for
  m.sourcesRow.content = content
  m.sourcesRow.jumpToRowItem = [0, 0]
  m.sourcesRow.SetFocus(true)
end sub

sub playSource(url as string)
  if url = invalid or url = ""
    showError("Selected source has no URL.")
    return
  end if

  item = CreateObject("roSGNode", "ContentNode")
  item.streamFormat = "mp4"
  item.url = url
  m.videoPlayer.content = item
  m.videoPlayer.control = "play"
  m.videoPlayer.visible = true
  m.videoPlayer.SetFocus(true)
end sub

sub togglePlayback()
  if m.videoPlayer.state = "playing"
    m.videoPlayer.control = "pause"
  else
    m.videoPlayer.control = "resume"
  end if
end sub

sub stopPlayback()
  m.videoPlayer.control = "stop"
  m.videoPlayer.visible = false
  m.sourcesRow.SetFocus(true)
end sub

sub updateSearchText()
  if m.searchText = ""
    m.searchLabel.text = "Type to search..."
  else
    m.searchLabel.text = m.searchText
  end if
end sub

sub updateSearchType()
  m.searchTypeLabel.text = m.searchTypes[m.searchTypeIndex]
end sub

sub updateSettingsLabels()
  m.addonUrlText.text = m.settings.addonUrl
  m.ratingText.text = m.settings.maxRating
  if m.settings.apiKey = ""
    m.apiKeyText.text = "(required)"
  else
    m.apiKeyText.text = m.settings.apiKey
  end if
end sub

sub saveSettings()
  sec = CreateObject("roRegistrySection", "StreamFinder")
  sec.Write("addonUrl", m.settings.addonUrl)
  sec.Write("maxRating", m.settings.maxRating)
  sec.Write("apiKey", m.settings.apiKey)
  sec.Flush()
end sub

sub showError(message as string)
  m.errorDialog.message = message
  m.errorDialog.visible = true
end sub

function firstNonEmpty(value as dynamic, fallback as string) as string
  if value = invalid then return fallback
  if type(value) <> "roString" and type(value) <> "String" then return value
  if Trim(value) = "" then return fallback
  return value
end function

function indexOf(items as object, target as string) as integer
  for i = 0 to items.Count() - 1
    if items[i] = target then return i
  end for
  return 0
end function

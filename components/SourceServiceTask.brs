sub init()
  m.top.functionName = "run"
end sub

sub run()
  m.top.errorText = ""
  if m.top.action = "search"
    m.top.results = { items: searchTitles(m.top.query, m.top.mediaType, m.top.apiKey, m.top.maxRating) }
  else if m.top.action = "sources"
    m.top.results = { streams: fetchSources(m.top.addonUrl, m.top.mediaType, m.top.imdbId) }
  else if m.top.action = "details"
    m.top.results = { detail: fetchDetail(m.top.imdbId, m.top.apiKey) }
  end if
end sub

function searchTitles(query as string, mediaType as string, apiKey as string, maxRating as string) as object
  if query = "" then return []
  if apiKey = "" then
    m.top.errorText = "Set an OMDb API key in settings."
    return []
  end if

  transfer = CreateObject("roUrlTransfer")
  safeQuery = query.Replace(" ", "+")
  url = "https://www.omdbapi.com/?apikey=" + apiKey + "&type=" + mediaType + "&s=" + safeQuery
  transfer.SetUrl(url)
  response = transfer.GetToString()
  if response = invalid then
    m.top.errorText = "Could not reach lookup service."
    return []
  end if

  json = ParseJson(response)
  if json = invalid or json.Response <> "True" then
    m.top.errorText = "No titles found."
    return []
  end if

  filtered = []
  for each item in json.Search
    detail = fetchDetail(item.imdbID, apiKey)
    rated = "NR"
    if detail <> invalid and detail.Rated <> invalid then rated = detail.Rated
    if ratingAllowed(rated, maxRating)
      filtered.Push({
        title: item.Title,
        year: item.Year,
        imdbID: item.imdbID,
        type: item.Type,
        poster: normalizePoster(item.Poster),
        plot: detail.Plot,
        rated: rated
      })
    end if
  end for

  return filtered
end function

function fetchDetail(imdbId as string, apiKey as string) as dynamic
  if imdbId = "" or apiKey = "" then return invalid
  transfer = CreateObject("roUrlTransfer")
  transfer.SetUrl("https://www.omdbapi.com/?apikey=" + apiKey + "&i=" + imdbId + "&plot=full")
  response = transfer.GetToString()
  if response = invalid then return invalid
  json = ParseJson(response)
  return json
end function

function fetchSources(addonUrl as string, mediaType as string, imdbId as string) as object
  if addonUrl = "" or imdbId = "" then return []
  transfer = CreateObject("roUrlTransfer")
  transfer.SetUrl(addonUrl + "/stream/" + mediaType + "/" + imdbId + ".json")
  response = transfer.GetToString()
  if response = invalid then
    m.top.errorText = "Could not load sources from addon server."
    return []
  end if

  json = ParseJson(response)
  if json = invalid or json.streams = invalid then
    m.top.errorText = "No sources available."
    return []
  end if

  rows = []
  for each stream in json.streams
    titleText = stream.name
    subtitleText = "Press OK to play"
    if stream.title <> invalid and stream.title <> "" then
      titleText = stream.title
    end if
    rows.Push({ title: titleText, subtitle: subtitleText, url: stream.url })
  end for
  return rows
end function

function normalizePoster(uri as dynamic) as string
  if uri = invalid or uri = "N/A" then
    return "https://via.placeholder.com/220x300?text=No+Image"
  end if
  return uri
end function

function ratingAllowed(actual as string, maxRating as string) as boolean
  ladder = {
    "G": 1, "TV-G": 1,
    "PG": 2, "TV-PG": 2,
    "PG-13": 3, "TV-14": 3,
    "R": 4, "TV-MA": 4,
    "NC-17": 5
  }

  a = ladder.Lookup(actual)
  m = ladder.Lookup(maxRating)
  if m = invalid then m = 3
  if a = invalid then return true
  return a <= m
end function

sub init()
  m.poster = m.top.findNode("poster")
  m.title = m.top.findNode("title")
end sub

sub updateContent()
  content = m.top.itemContent
  if content = invalid then return
  m.poster.uri = content.poster
  m.title.text = content.title
end sub

sub init()
  m.title = m.top.findNode("title")
  m.sub = m.top.findNode("sub")
end sub

sub updateContent()
  content = m.top.itemContent
  if content = invalid then return
  m.title.text = content.title
  m.sub.text = content.subtitle
end sub

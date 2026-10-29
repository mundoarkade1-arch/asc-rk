' ===== Util.brs: funcoes comuns do ASC RK =====
function RegGet(key as String) as String
  s = CreateObject("roRegistrySection", "ascrk")
  if s.Exists(key) then return s.Read(key)
  return ""
end function

sub RegSet(key as String, value as String)
  s = CreateObject("roRegistrySection", "ascrk")
  s.Write(key, value)
  s.Flush()
end sub

sub RegDel(key as String)
  s = CreateObject("roRegistrySection", "ascrk")
  s.Delete(key)
  s.Flush()
end sub

function Sx(v as Dynamic) as String
  if v = invalid then return ""
  t = type(v)
  if t = "roString" or t = "String" then return v
  if t = "roInt" or t = "Integer" or t = "roFloat" or t = "Float" or t = "LongInteger" or t = "roLongInteger" or t = "Double" or t = "roDouble" then return v.ToStr()
  return ""
end function

function Pad2(n as Integer) as String
  if n < 10 then return "0" + n.ToStr()
  return n.ToStr()
end function

function FmtTime(sec as Dynamic) as String
  t = Int(sec)
  h = t \ 3600
  mi = (t mod 3600) \ 60
  s = t mod 60
  if h > 0 then return h.ToStr() + ":" + Pad2(mi) + ":" + Pad2(s)
  return Pad2(mi) + ":" + Pad2(s)
end function

function ClockStr(ts as Integer) as String
  d = CreateObject("roDateTime")
  d.FromSeconds(ts)
  d.ToLocalTime()
  return Pad2(d.GetHours()) + ":" + Pad2(d.GetMinutes())
end function

function JoinStr(arr as Object, sep as String) as String
  out = ""
  for i = 0 to arr.Count() - 1
    if i > 0 then out = out + sep
    out = out + arr[i]
  end for
  return out
end function

function Norm(s as String) as String
  pairs = ["á","a","à","a","â","a","ã","a","ä","a","é","e","è","e","ê","e","ë","e","í","i","ì","i","î","i","ï","i","ó","o","ò","o","ô","o","õ","o","ö","o","ú","u","ù","u","û","u","ü","u","ç","c","ñ","n","Á","a","À","a","Â","a","Ã","a","É","e","Ê","e","Í","i","Ó","o","Ô","o","Õ","o","Ú","u","Ç","c"]
  for i = 0 to pairs.Count() - 1 step 2
    if Instr(1, s, pairs[i]) > 0 then s = s.Replace(pairs[i], pairs[i + 1])
  end for
  return LCase(s)
end function

function B64(s as String) as String
  if s = "" then return ""
  b = CreateObject("roByteArray")
  if b.FromBase64String(s) then return b.ToAsciiString()
  return s
end function

function UrlEnc(s as String) as String
  out = ""
  for i = 1 to Len(s)
    c = Mid(s, i, 1)
    a = Asc(c)
    if (a >= 48 and a <= 57) or (a >= 65 and a <= 90) or (a >= 97 and a <= 122) or c = "-" or c = "_" or c = "." or c = "~"
      out = out + c
    else
      h = UCase(StrI(a, 16))
      if Len(h) < 2 then h = "0" + h
      out = out + "%" + h
    end if
  end for
  return out
end function

function Api(act as String, extra = "" as String) as String
  s = m.global.session
  u = s.server + "/player_api.php?username=" + UrlEnc(s.user) + "&password=" + UrlEnc(s.pass)
  if act <> "" then u = u + "&action=" + act
  return u + extra
end function

function StreamUrl(kind as String, id as String, ext as String) as String
  s = m.global.session
  return s.server + "/" + kind + "/" + UrlEnc(s.user) + "/" + UrlEnc(s.pass) + "/" + id + "." + ext
end function

' So pode ser usada dentro de Task (thread de rede)
function HttpGet(url as String) as Dynamic
  x = CreateObject("roUrlTransfer")
  port = CreateObject("roMessagePort")
  x.SetPort(port)
  x.SetUrl(url)
  x.SetCertificatesFile("common:/certs/ca-bundle.crt")
  x.InitClientCertificates()
  x.EnableEncodings(true)
  x.AddHeader("User-Agent", "ASC-RK/1.0")
  if not x.AsyncGetToString() then return invalid
  msg = wait(20000, port)
  if type(msg) = "roUrlEvent"
    if msg.GetResponseCode() = 200 then return ParseJson(msg.GetString())
  else
    x.AsyncCancel()
  end if
  return invalid
end function

sub Fetch(url as String, tag as String, cb as String)
  if m.taskMap = invalid then m.taskMap = {}
  t = CreateObject("roSGNode", "ApiTask")
  t.tag = tag
  t.url = url
  t.observeField("result", cb)
  m.taskMap[tag] = t
  t.control = "RUN"
end sub

function MakeItem(kind as String) as Object
  return { kind: kind, sid: "", title: "", hdposterurl: "", ext: "", plot: "", rating: "", pos: 0, dur: 0, key: "", series: "", season: 0, epnum: 0, direct: false }
end function

function MakeInfo(kind as String, sid as String, title as String, cover as String, ext as String) as Object
  return { kind: kind, sid: sid, title: title, cover: cover, ext: ext, plot: "", rating: "", pos: 0, dur: 0, key: "", series: "", season: 0, epnum: 0, direct: true }
end function

function ItemInfo(n as Object) as Object
  return { kind: n.kind, sid: n.sid, title: n.title, cover: n.hdposterurl, ext: n.ext, plot: n.plot, rating: n.rating, pos: n.pos, dur: n.dur, key: n.key, series: n.series, season: n.season, epnum: n.epnum, direct: n.direct }
end function

function BuildList(arr as Object, kind as String) as Object
  list = []
  if type(arr) <> "roArray" then return list
  n = arr.Count()
  if n > 1500 then n = 1500
  for i = 0 to n - 1
    it = arr[i]
    if type(it) = "roAssociativeArray"
      o = MakeItem(kind)
      if kind = "live"
        o.sid = Sx(it.stream_id)
        o.title = Sx(it.name)
        o.hdposterurl = Sx(it.stream_icon)
      else if kind = "movie"
        o.sid = Sx(it.stream_id)
        o.title = Sx(it.name)
        o.hdposterurl = Sx(it.stream_icon)
        o.ext = Sx(it.container_extension)
        o.rating = Sx(it.rating)
      else
        o.sid = Sx(it.series_id)
        o.title = Sx(it.name)
        o.hdposterurl = Sx(it.cover)
        o.plot = Sx(it.plot)
        o.rating = Sx(it.rating)
      end if
      list.Push(o)
    end if
  end for
  return list
end function

function BuildNode(list as Object) as Object
  n = CreateObject("roSGNode", "ContentNode")
  n.Update({ children: list }, true)
  return n
end function

' ----- Continuar assistindo (guardado no registro da TV) -----
function LoadProgress() as Object
  j = ParseJson(RegGet("progress"))
  if j = invalid then return {}
  return j
end function

sub SaveProgressItem(item as Object)
  p = LoadProgress()
  if item.series <> ""
    for each k in p.Keys()
      v = p[k]
      if type(v) = "roAssociativeArray"
        if v.series = item.series and k <> item.key then p.Delete(k)
      end if
    end for
  end if
  p[item.key] = item
  keys = p.Keys()
  if keys.Count() > 40
    oldest = keys[0]
    ot = 9999999999
    for each k in keys
      t = p[k].t
      if t < ot
        ot = t
        oldest = k
      end if
    end for
    p.Delete(oldest)
  end if
  RegSet("progress", FormatJson(p))
end sub

sub RemoveProgress(key as String)
  p = LoadProgress()
  if p.DoesExist(key)
    p.Delete(key)
    RegSet("progress", FormatJson(p))
  end if
end sub

function FindSeriesProgress(sid as String) as Dynamic
  p = LoadProgress()
  for each k in p.Keys()
    v = p[k]
    if type(v) = "roAssociativeArray"
      if v.series = sid then return v
    end if
  end for
  return invalid
end function

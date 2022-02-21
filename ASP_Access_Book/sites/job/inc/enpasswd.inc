<% function encrypt(ecode)
		Dim texts
		dim i
		for i=1 to len(ecode)
			texts=texts & chr(asc(mid(ecode,i,1))+i)
		next
      encrypt = texts
      end function 
	function decrypt(dcode)	
  		dim texts
		dim i
		for i=1 to len(dcode)
			texts=texts & chr(asc(mid(dcode,i,1))-i)
		next
		decrypt=texts
	end function
	function	mistake(preString)
      Dim texts
      Dim seed
      Dim i,length
      prestring = trim(preString)
      length = len(preString)
      seed = length
      Randomize(length)
      texts = ""
      for i = 1 to length
            seed = int(94*rnd(-asc(mid(preString,i,1))-seed*asc(right(prestring,1)))+32)
            texts = texts & chr(seed) & chr(int(94*rnd(-seed)+32))
      next
	dim dist
	dist=""
	for i = 1 to len(texts)
	  if mid(texts,i,1)<>"'"  then
		      dist=dist+mid(texts,i,1)
      	end if
    next 
    mistake = dist
	end function
	  %>

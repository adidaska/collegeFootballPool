<%	'**************************************************************************
	'* Common code for encrypting/decrypting data (uses Alleged RC4 cipher).  *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Encrypts the given string using the global encryption key. The string
	' returned is a hex representation.
	'--------------------------------------------------------------------------
	function Encrypt(str)

		Encrypt = CharToHex(Arc4Transform(str, DATA_ENCRYPTION_KEY))

	end function

	'--------------------------------------------------------------------------
	' Decrypts a string encrypted via Encrypt().
	'--------------------------------------------------------------------------
	function Decrypt(str)

		Decrypt = Arc4Transform(HexToChar(str), DATA_ENCRYPTION_KEY)

	end function

	'--------------------------------------------------------------------------
	' Returns a hex representation of the given string.
	'--------------------------------------------------------------------------
	function CharToHex(str)

		dim i

		CharToHex = ""
		for i = 1 to Len(str)
			CharToHex = CharToHex & Right("0" & Hex(Asc(Mid(str, i, 1))), 2)
		next

	end function

	'--------------------------------------------------------------------------
	' Returns a string given a hex representation.
	'--------------------------------------------------------------------------
	function HexToChar(str)

		dim i, s

		HexToChar = ""

		'Exit if a null value was passed.
		if IsNull(str) then
			exit function
		end if

		for i = 1 to Len(str) step 2
			s = "&h" & Mid(str, i, 2)

			'Return and empty string if conversion fails.
			if not IsNumeric(s) then
				HexToChar = ""
				exit function
			end if

			HexToChar = HexToChar & Chr(CInt(s))
		next

	end function

	'Define the permutation array.
	dim Arc4S(255)

	'--------------------------------------------------------------------------
	' Initializes the permutation array.
	'--------------------------------------------------------------------------
	sub Arc4Initialize(key)

		dim i, j, keyLen, temp

		for i = 0 to 255
			Arc4S(i) = i
		next
		keyLen = Len(key)
		j = 0
		for i = 0 to 255
			j = (j + Arc4S(i) + Asc(Mid(key, (i mod keyLen) + 1, 1))) mod 256
			temp = Arc4S(i)
			Arc4S(i) = Arc4S(j)
			Arc4S(j) = temp
		next

	end sub

	'--------------------------------------------------------------------------
	' Performs the encryption/decryption.
	'--------------------------------------------------------------------------
	function Arc4Transform(str, key)

		dim i, j, n, temp, k, x

		Arc4Transform = ""

		'Exit if a null value was passed.
		if IsNull(str) then
			exit function
		end if

		call Arc4Initialize(key)
		i = 0
		j = 0

		for n = 1 to Len(str)
			i = (i + 1) mod 256
			j = (j + Arc4S(i)) mod 256
			temp = Arc4S(i)
			Arc4S(i) = Arc4S(j)
			Arc4S(j) = temp
			k = Arc4S((Arc4S(i) + Arc4S(j)) mod 256)
			x = Asc(Mid(str, n, 1)) xor k
			Arc4Transform = Arc4Transform & Chr(x)
		next

	end function %>

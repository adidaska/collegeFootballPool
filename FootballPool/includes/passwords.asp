<%	'**************************************************************************
	'* Common code for handling user passwords.                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns an eight-character random salt string. Used for generating the
	' hash for passwords.
	'--------------------------------------------------------------------------
	function CreateSalt()

		const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		dim i, n

		CreateSalt = ""
		Randomize
		for i = 1 to 8
			n = Int(Len(chars) * Rnd()) + 1
			CreateSalt = CreateSalt & Mid(chars, n, 1)
		next

	end function

	'--------------------------------------------------------------------------
	' Returns an eight-character random character string. Used for generating
	' temporary passwords.
	'--------------------------------------------------------------------------

	function CreatePassword()

		'Define a set of random alphanumerics.
		'Note that we omit letters and numbers that can be easily confused,
		'such as the letter 'l' and the number '1'.
		const chars = "abcdefghijkmnpqrstuvwxyz23456789"
		dim i, n

		CreatePassword = ""
		Randomize
		for i = 1 to 8
			n = Int(Len(chars) * Rnd()) + 1
			CreatePassword = CreatePassword & Mid(chars, n, 1)
		next

	end function

	'--------------------------------------------------------------------------
	' Returns a 40-character hash value for the given string. Used for storing
	' passwords.
	'--------------------------------------------------------------------------
	function Hash(str)

		'Call the JScript function to calculate the SHA-1 message digest for
		'the string.
		Hash = SecureHash(str)

	end function %>

<SCRIPT LANGUAGE="JScript" RUNAT="Server">
	//*************************************************************************
	// Password hashing functions (note change in scripting language).
	//*************************************************************************

	//-------------------------------------------------------------------------
	// Calculates the message digest for the specified string using the SHA-1
	// algorithm.
	//-------------------------------------------------------------------------
	function SecureHash(s)
	{
		var bitLen, n, w;
		var k, buffer, f;
		var h0, h1, h2, h3, h4;
		var a, b, c, d, e, temp;
		var i, j;

		// Pad the input message per the SHA-1 specification.
		bitLen = s.length * 8;
		n = 512 - (bitLen % 512);
		if (n <= 65)
			n += 512;
		s += String.fromCharCode(0x80);
		n -= 8;
		for (i = 0; i < n / 8; i++)
			s += String.fromCharCode(0x00);

		// Convert the padded message to an array of 32-bit integers, note
		// conversion to little endian.
		w = new Array();
		for (i = 0; i < s.length; i += 4)
		{
			n = 0;
			for (j = 0; j < 4; j++)
				n = (n << 8) + s.charCodeAt(i + j);
			w[w.length] = n;
		}

		// Set the last word to the original bit length of the message.
		w[w.length - 1] = bitLen;

		// Initialize the hash values, constants and buffer.
		h0 = 0x67452301;
		h1 = 0xEFCDAB89;
		h2 = 0x98BADCFE;
		h3 = 0x10325476;
		h4 = 0xC3D2E1F0;

		k = new Array(80);
		for (i = 0; i < 80; i++)
			if (i < 20)
				k[i] = 0x5A827999;
			else if (i < 40)
				k[i] = 0x6ED9EBA1;
			else if (i < 60)
				k[i] = 0x8F1BBCDC;
			else
				k[i] = 0xCA62C1D6;
		buffer = new Array(80);

		// Process the word array in 512-bit (16-word) blocks.
		n = w.length / 16;
		for (i = 0; i < n; i ++)
		{
			// Initialize the 80-word buffer using the current block.
			for (j = 0; j < 80; j++)
				if (j < 16)
					buffer[j] = w[i * 16 + j];
				else
					buffer[j] = Rol32(buffer[j - 3] ^ buffer[j - 8] ^ buffer[j - 14] ^ buffer[j - 16], 1);


			// Hash the block.
			a = h0; b = h1; c = h2; d = h3; e = h4;
			for (j = 0; j < 80; j++)
			{
				temp = Rol32(a, 5);
				if (j < 20)
					f = (b & c) | ((~b) & d);
				else if (j < 40)
					f = b ^ c ^ d;
				else if (j < 60)
					f = (b & c) | (b & d) | (c & d);
				else
					f = b ^ c ^ d;
				temp = Add32(temp, f);
				temp = Add32(temp, e);
				temp = Add32(temp, buffer[j]);
				temp = Add32(temp, k[j]);
				e = d; d = c; c = Rol32(b, 30); b = a; a = temp;
			}

			// Update the hash values.
			h0 = Add32(h0, a);
			h1 = Add32(h1, b);
			h2 = Add32(h2, c);
			h3 = Add32(h3, d);
			h4 = Add32(h4, e);
		}

     	// Format the hash values and return them as the message digest.
		return Hex32(h0) + Hex32(h1) + Hex32(h2) + Hex32(h3) + Hex32(h4);
	}

	//-------------------------------------------------------------------------
	// Helper functions for the hash function.
	//-------------------------------------------------------------------------
	function Rol32(x, n)
	{
		// Left circular shift for a 32-bit integer.
		return (x << n) | (x >>> (32 - n));
  	}

	function Add32(x, y)
	{
		// Add two 32-bit integers, wrapping at 2^32.
		return ((x & 0x7FFFFFFF) + (y & 0x7FFFFFFF)) ^ (x & 0x80000000) ^ (y & 0x80000000);
	}

	function Hex32(n)
	{
		var hexDigits = "0123456789ABCDEF";
		var i, s;

		// Format a 32-bit integer as a hexadecimal string.
		s = "";
		for (i = 7; i >= 0; i--)
			s += hexDigits.charAt((n >> (i * 4)) & 0x0F);

		return s;
	}
</SCRIPT>

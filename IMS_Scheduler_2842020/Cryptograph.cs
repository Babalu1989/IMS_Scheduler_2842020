using System;
using System.Collections;
using System.Text;
using System.Configuration;

public class Cryptograph
{
    private const string PASSWORD_ENCRYPTION_KEY = "o8??^am(*)";
    private const string SPECIAL_PASSWORD_ENCRYPTION_KEY = "aishwaryarai007";
    private long[] x1a0;
    private long[] cle;
    private long x1a2;
    private long inter;
    private long res, ax, bx, cx, dx, si, tmp, c, d, e;
    private long I;
    private string Key;
    public Cryptograph()
    {
        x1a0 = new long[10];
        cle = new long[18];
    }
    private void Assemble()
    {
        x1a0[0] = ((cle[1] * 256) + cle[2]) % 65536;
        code();
        inter = res;

        x1a0[1] = x1a0[0] ^ ((cle[3] * 256) + cle[4]);
        code();
        inter = inter ^ res;


        x1a0[2] = x1a0[1] ^ ((cle[5] * 256) + cle[6]);
        code();
        inter = inter ^ res;

        x1a0[3] = x1a0[2] ^ ((cle[7] * 256) + cle[8]);
        code();
        inter = inter ^ res;

        x1a0[4] = x1a0[3] ^ ((cle[9] * 256) + cle[10]);
        code();
        inter = inter ^ res;

        x1a0[5] = x1a0[4] ^ ((cle[11] * 256) + cle[12]);
        code();
        inter = inter ^ res;

        x1a0[6] = x1a0[5] ^ ((cle[13] * 256) + cle[14]);
        code();
        inter = inter ^ res;

        x1a0[7] = x1a0[6] ^ ((cle[15] * 256) + cle[16]);
        code();
        inter = inter ^ res;
        I = 0;
    }
    private void code()
    {
        dx = (x1a2 + I) % 65536;
        ax = x1a0[I];
        cx = 0x15A;
        bx = 0x4E35;

        tmp = ax;
        ax = si;
        si = tmp;

        tmp = ax;
        ax = dx;
        dx = tmp;

        if (ax != 0)
        {
            ax = (ax * bx) % 65536;
        }

        tmp = ax;
        ax = cx;
        cx = tmp;

        if (ax != 0)
        {
            ax = (ax * si) % 65536;
            cx = (ax + cx) % 65536;
        }

        tmp = ax;
        ax = si;
        si = tmp;
        ax = (ax * bx) % 65536;
        dx = (cx + dx) % 65536;

        ax = ax + 1;

        x1a2 = dx;
        x1a0[I] = ax;

        res = ax ^ dx;
        I = I + 1;
    }

    public string Encrypt(string EncryptText, string PW_KEY)
    {
        string EncryptedText, champ1;
        long compte, cfc, cfd;
        Int32 fois;
        char temp;
        EncryptedText = "";
        Key = PW_KEY;
        si = 0;
        x1a2 = 0;
        I = 0;

        for (fois = 0; fois <= 17; fois++)
            cle[fois] = 0;
        champ1 = Key;

        for (fois = 0; fois < champ1.ToString().Length; )
        {
            temp = char.Parse(champ1.Substring(fois, 1));
            cle[++fois] = (long)temp;
        }

        champ1 = EncryptText;
        for (fois = 0; fois < champ1.ToString().Length; fois++)
        {

            c = (long)char.Parse(champ1.Substring(fois, 1));
            Assemble();
            if (inter > 65535)
                inter = inter - 65536;
            cfc = (((inter * 256) / 256) - (inter % 256)) / 256;
            cfd = inter % 256;
            for (compte = 1; compte < 17; compte++)
                cle[compte] = cle[compte] ^ c;

            c = c ^ (cfc ^ cfd);
            d = (((c * 16) / 16) - (c % 16)) / 16;
            e = c % 16;

            EncryptedText = EncryptedText + (char)(0x61 + d);
            EncryptedText = EncryptedText + (char)(0x61 + e);
        }
        //}

        return (EncryptedText);
    }
    public string Encrypt(string EncryptText)
    {
        string PW_KEY = SPECIAL_PASSWORD_ENCRYPTION_KEY;
        string EncryptedText, champ1;
        long compte, cfc, cfd;
        Int32 fois;
        char temp;
        EncryptedText = "";
        Key = PW_KEY;
        si = 0;
        x1a2 = 0;
        I = 0;

        for (fois = 0; fois <= 17; fois++)
            cle[fois] = 0;
        champ1 = Key;

        for (fois = 0; fois < champ1.ToString().Length; )
        {
            temp = char.Parse(champ1.Substring(fois, 1));
            cle[++fois] = (long)temp;
        }

        champ1 = EncryptText;
        for (fois = 0; fois < champ1.ToString().Length; fois++)
        {

            c = (long)char.Parse(champ1.Substring(fois, 1));
            Assemble();
            if (inter > 65535)
                inter = inter - 65536;
            cfc = (((inter * 256) / 256) - (inter % 256)) / 256;
            cfd = inter % 256;
            for (compte = 1; compte < 17; compte++)
                cle[compte] = cle[compte] ^ c;

            c = c ^ (cfc ^ cfd);
            d = (((c * 16) / 16) - (c % 16)) / 16;
            e = c % 16;

            EncryptedText = EncryptedText + (char)(0x61 + d);
            EncryptedText = EncryptedText + (char)(0x61 + e);
        }

        return (EncryptedText);
    }

    public string Decrypt(string DecryptText, string PW_KEY)
    {
        string result, champ1;
        long compte;
        Int32 fois;

        long cfc, cfd;
        result = "";
        char temp;
        Key = PW_KEY;
        si = 0;
        x1a2 = 0;
        I = 0;

        for (fois = 0; fois <= 17; fois++)
            cle[fois] = 0;
        champ1 = Key;

        for (fois = 0; fois < champ1.ToString().Length; )
        {
            temp = char.Parse(champ1.Substring(fois, 1));
            cle[++fois] = (long)temp;
        }
        champ1 = DecryptText;
        for (fois = 0; fois < champ1.ToString().Length; fois++)
        {
            d = (long)char.Parse(champ1.Substring(fois, 1));
            if ((d - 0x61) >= 0)
            {
                d = d - 0x61; 
                if ((d >= 0) & (d <= 15))
                    d = d * 16;
            }
            if (fois != champ1.ToString().Length)
                fois++;
            e = (long)char.Parse(champ1.Substring(fois, 1));
            if ((e - 0x61) >= 0)
            {
                e = e - 0x61; 
                if ((e >= 0) & (e <= 15))
                    c = d + e;
            }
            Assemble();

            if (inter > 65535)
                inter = inter - 65536;
            cfc = (((inter * 256) / 256) - (inter % 256)) / 256;
            cfd = inter % 256;

            c = c ^ (cfc ^ cfd);
            for (compte = 1; compte <= 16; compte++)
                cle[compte] = cle[compte] ^ c;
            result = result + (char)(c);
        }
        return (result);
    }
    public string Decrypt(string DecryptText)
    {
        string PW_KEY = SPECIAL_PASSWORD_ENCRYPTION_KEY;
        string result, champ1;
        long compte;
        Int32 fois;

        long cfc, cfd;
        result = "";
        char temp;
        Key = PW_KEY;
        si = 0;
        x1a2 = 0;
        I = 0;

        for (fois = 0; fois <= 17; fois++)
            cle[fois] = 0;
        champ1 = Key;

        for (fois = 0; fois < champ1.ToString().Length; )
        {
            temp = char.Parse(champ1.Substring(fois, 1));
            cle[++fois] = (long)temp;
        }
        champ1 = DecryptText;
        for (fois = 0; fois < champ1.ToString().Length; fois++)
        {
            d = (long)char.Parse(champ1.Substring(fois, 1));
            if ((d - 0x61) >= 0)
            {
                d = d - 0x61;  
                if ((d >= 0) & (d <= 15))
                    d = d * 16;
            }
            if (fois != champ1.ToString().Length)
                fois++;
            e = (long)char.Parse(champ1.Substring(fois, 1));
            if ((e - 0x61) >= 0)
            {
                e = e - 0x61; 
                if ((e >= 0) & (e <= 15))
                    c = d + e;
            }
            Assemble();

            if (inter > 65535)
                inter = inter - 65536;
            cfc = (((inter * 256) / 256) - (inter % 256)) / 256;
            cfd = inter % 256;

            c = c ^ (cfc ^ cfd);
            for (compte = 1; compte <= 16; compte++)
                cle[compte] = cle[compte] ^ c;
            result = result + (char)(c);
        }
        return (result);
    }
}

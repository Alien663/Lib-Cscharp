using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Alien.Common.Config;

public class AESCrypto : IDisposable
{
    private bool disposed = false;
    private byte[] once;
    private byte[] _key;

    public AESCrypto(string key = "Alien19BxNd4BMFGSjROulGESLTVeZjC", string iv = "8b85a31a084t")
    {
        _key = Encoding.UTF8.GetBytes(key);
        once = Encoding.UTF8.GetBytes(iv);
    }

    public string Encrypt(string plainText)
    {
        try
        {
            byte[] plaintextBytes = Encoding.UTF8.GetBytes(plainText);
            byte[] ciphertext = new byte[plaintextBytes.Length];
            byte[] tag = new byte[AesGcm.TagByteSizes.MaxSize];
            using (AesGcm aesGcm = new AesGcm(_key, 16))
            {
                aesGcm.Encrypt(once, plaintextBytes, ciphertext, tag);
                return Convert.ToBase64String(once.Concat(ciphertext).Concat(tag).ToArray());
            }
        }
        catch (Exception ex)
        {
            throw new CryptographicException("Encryption failed", ex);
        }
    }

    public string Decrypt(string cipherText)
    {
        try
        {
            byte[] fullCipher = Convert.FromBase64String(cipherText);
            byte[] tag = fullCipher.Skip(fullCipher.Length - AesGcm.TagByteSizes.MaxSize).ToArray();
            byte[] ciphertext = fullCipher.Skip(AesGcm.NonceByteSizes.MaxSize).Take(fullCipher.Length - AesGcm.NonceByteSizes.MaxSize - AesGcm.TagByteSizes.MaxSize).ToArray();
            byte[] plaintextBytes = new byte[ciphertext.Length];

            using (AesGcm aesGcm = new AesGcm(_key, 16))
            {
                aesGcm.Decrypt(once, ciphertext, tag, plaintextBytes);
            }

            return Encoding.UTF8.GetString(plaintextBytes);
        }
        catch (Exception ex)
        {
            throw new CryptographicException("Decryption failed", ex);
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private void Dispose(bool disposing)
    {
        if (!disposed)
        {
            if (disposing)
            {
                // Dispose managed resources.
                Array.Clear(_key, 0, _key.Length);
                Array.Clear(once, 0, once.Length);
            }
            // Dispose unmanaged resources.
            disposed = true;
        }
    }
}

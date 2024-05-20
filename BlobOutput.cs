using System.IO;

namespace ReadExcel
{
    public class BlobOutput
    {
        public string BlobName { get; set; }
        public Stream BlobContent { get; set; }
    }
}
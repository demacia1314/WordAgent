import yauzl from 'yauzl';

export function validateDocxArchive(buffer: Buffer): Promise<void> {
  return new Promise((resolve, reject) => {
    yauzl.fromBuffer(buffer, { lazyEntries: true, validateEntrySizes: true }, (error, archive) => {
      if (error || !archive) return reject(new Error('文档压缩包无法读取'));
      let total = 0;
      let count = 0;
      let documentFound = false;
      let contentTypesFound = false;
      let failed = false;
      const fail = (message: string) => {
        if (failed) return;
        failed = true;
        archive.close();
        reject(new Error(message));
      };
      archive.on('error', () => fail('文档压缩包已损坏'));
      archive.on('entry', (entry) => {
        total += entry.uncompressedSize;
        count++;
        if (total > 48 * 1024 * 1024 || count > 3000 || entry.uncompressedSize > 24 * 1024 * 1024)
          return fail('文档解压后过大，请拆分文件');
        if (entry.isEncrypted()) return fail('不支持加密文档');
        if (entry.fileName === 'word/document.xml') documentFound = true;
        if (entry.fileName === '[Content_Types].xml') contentTypesFound = true;
        archive.readEntry();
      });
      archive.on('end', () => {
        if (!failed) {
          if (documentFound && contentTypesFound) resolve();
          else reject(new Error('不是有效的 Word 文档'));
        }
      });
      archive.readEntry();
    });
  });
}

#ifndef  DCG_SQLITE_CRYPT_FUNC_
#define  DCG_SQLITE_CRYPT_FUNC_

//#ifdef SQLITE_HAS_CODEC

typedef unsigned char BYTE;
// 加密结构
#define CRYPT_OFFSET 8
typedef struct _CryptBlock {

	BYTE* ReadKey; // 读数据库和写入事务的密钥
	BYTE* WriteKey; // 写入数据库的密钥
	int PageSize; // 页的大小
	BYTE* Data;

} CryptBlock, *LPCryptBlock;

#ifndef DB_KEY_LENGTH_BYTE        // 密钥长度
#define DB_KEY_LENGTH_BYTE   16   // 密钥长度
#endif

#ifndef DB_KEY_PADDING            // 密钥位数不足时补充的字符
#define DB_KEY_PADDING       0x33 // 密钥位数不足时补充的字符
#endif

// 加密函数
extern int sqlite3_encrypt(unsigned char * pData, unsigned int data_len, unsigned char * key, unsigned int len_of_key);
 
// 解密函数
extern int sqlite3_dencrypt(unsigned char * pData, unsigned int data_len, unsigned char * key, unsigned int len_of_key);

// ===
// 被sqlite3调用的加/解密函数
extern void* sqlite3Codec(void *pArg, void *data, int nPageNum, int nMode);

// 释放而加/密内存的回调
extern void sqlite3CodecFree(void* arg);

// 验证密码接口
extern unsigned char* DeriveKey(const void *pKey, int nKeyLen);

// 创建或更新一个页的加密算法索引.此函数会申请缓冲区.
extern LPCryptBlock CreateCryptBlock(unsigned char* hKey, int nPageSize, LPCryptBlock pExisting);

//#endif // SQLITE_HAS_CODEC
#endif



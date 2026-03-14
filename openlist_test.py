from openlist import Client
import asyncio

openlist_url="https://openlist.tonyv2.top"
openlist_username="Tony_V2"   
openlist_password="Tony090721"    
openlist_target_folder="/阿里云盘/课件备份%20高二"

async def main():
    async with Client(openlist_url) as client:
        await client.login(openlist_username, openlist_password)
        user_info = await client.user.me()
        print(str(user_info))
        fs = client.fs
        files = await fs.listdir("/")
        for f in files:
            print(f"{f.name} - {'目录' if f.is_dir else '文件'}")

asyncio.run(main())
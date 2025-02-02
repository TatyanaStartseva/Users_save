import os
import psycopg2
from aiohttp import web
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()
HOST = os.getenv("HOST")
DATABASE = os.getenv("DATABASE")
USER = os.getenv("USERNAME_DB")
PASSWORD = os.getenv("PASSWORD_DB")


async def handle_chats_links(request):
    try:
        data = request.query
        urls = data.get("urls").split(",")
        conn = psycopg2.connect(
            host=HOST, database=DATABASE, user=USER, password=PASSWORD
        )
        cursor = conn.cursor()
        chat_ids = []
        for url in urls:
            cursor.execute(
                "SELECT chat_id FROM chats WHERE parent_link = %s OR children_link = %s",
                (url, url),
            )
            chat = cursor.fetchone()
            if chat:
                chat_ids.append(chat[0])

        chat_users = []
        user_ids_written = set()
        if chat_ids:
            for chat_id in chat_ids:
                cursor.execute("SELECT user_id FROM user_chat WHERE chat_id = %s", (chat_id,))
                users = cursor.fetchall()
                for user in users:
                    cursor.execute("SELECT * FROM users WHERE user_id = %s", (user,))
                    user_data = cursor.fetchall()
                    user_id = user_data[0][0]
                    if user_id not in user_ids_written:
                        chat_users.append(
                            (
                                user_data[0][0],
                                user_data[0][1],
                                user_data[0][2],
                                user_data[0][3],
                                user_data[0][4],
                                user_data[0][5],  # last_online
                                user_data[0][6],  # premium
                                user_data[0][7],  # phone
                                user_data[0][8],  # image
                            )
                        )
                        user_ids_written.add(user_id)
        else:
            return web.Response(status=404)

        wb = Workbook()
        ws = wb.active
        ws.append(
            [
                "user_id",
                "username",
                "bio",
                "first_name",
                "last_name",
                "last_online",
                "premium",
                "phone",
                "image",
            ]
        )
        for user in chat_users:
            user_data = [
                user[0],
                user[1],
                user[2],
                user[3],
                user[4],
                user[5].strftime("%Y-%m-%d %H:%M:%S") if user[5] is not None else "",
                "false" if user[6] == False else "true",
                "" if user[7] is None else user[7],
                "true" if user[8] == True else "false",
            ]
            ws.append(user_data)

        file_path = "chats_users.xlsx"
        wb.save(file_path)
        with open(file_path, "rb") as f:
            content = f.read()

        response = web.Response(
            body=content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        os.remove(file_path)
        return response
    except Exception as e:
        return web.Response(text=f"Error: {e}")


app = web.Application()
app.router.add_get("/chats_links", handle_chats_links)

web.run_app(app, port=80)

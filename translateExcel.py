from openai import OpenAI
import openpyxl
import datetime
import time
import asyncio

apiKey = input("Enter your OpenAI API key below\n")

client = OpenAI(api_key=apiKey)

fileName = input("Enter the name of the file you want to translate, including the extension\n")
    
workbook = openpyxl.load_workbook(fileName)
sheet = workbook.active
numRows = sheet.max_row
numCols = sheet.max_column
print("Number of rows:", numRows)
print("Number of columns:", numCols)

confirm = input("Are you sure you want to proceed? Enter 'yes' to continue.\n")
if confirm != "yes":
    exit()

def make_request(content: str):
    for attempt in range(5):
        try:
            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {
                        "role": "user",
                        "content": "The following are strings delimited by the # character. Translate each string into English, unless it is already in English. Return in the same format as the input. Do not make any notes. Here:\n" + content
                    }
                ],
                model="gpt-4o-mini"
            )
            return response
        except Exception:
            wait_time = 2 ** attempt
            print("rate limit error, retrying")
            time.sleep(wait_time)
    raise Exception("Failed to make request")

async def translateRow(i):
    loop = asyncio.get_event_loop()
    row = []
    for j in range(1, numCols + 1):
        cell = sheet.cell(i, j)
        row.append(cell.value if cell.value else "None")
    rowStr = "#".join(row)
    newRow = []
    tries = 0
    while len(newRow) != len(row) and tries < 5:
        response = await loop.run_in_executor(None, make_request, rowStr)
        newRowStr = response.choices[0].message.content
        newRow = newRowStr.split("#")
        tries += 1
    for j in range(1, numCols + 1):
        cell = sheet.cell(i, j)
        if newRow[j - 1] != "None":
            cell.value = newRow[j - 1]
    sheet.row_dimensions[i].height = None


async def main():
    startTime = datetime.datetime.now()

    tasks = [translateRow(i) for i in range(1, numRows + 1)]
    await asyncio.gather(*tasks) #unpacking operator *

    endTime = datetime.datetime.now()
    print("Time taken:", endTime - startTime)

    timeStr = str(endTime.year) + str(endTime.month) + str(endTime.day) + str(endTime.hour) + str(endTime.minute) + str(endTime.second)
    workbook.save(timeStr + fileName)

asyncio.run(main())
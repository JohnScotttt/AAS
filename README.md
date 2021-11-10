# AutoApiServer

## Purpose
* Renew MS E5 account.

## Explanation
* The project is based on ~~wangziyingwen/AutoApi~~.
* Most of the code is contributed by wangziyingwen.
* I rebuild it to make it run on server.
* **Renewal cannot be guaranteed**.
* **This method is for personal use only, please do not spread it**.

## Needed
* An **MS E5 account(not private account)**.
* A **server** or an environment that can run automatically **on a regular basis**.
* Python 3.x

## Usage
### MS E5
* I will not teach you how to get an MS E5 account, you need to get it by yourself online.
* If you have an MS E5 administrator account, next time we will use **it**, which ending with 'onmicrosoft.com'.
### Azure
* Open [Azure](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade), sign in the account.
* Click **+ New registration**.
    * Fill in **Name**, select the third item for the **supported account types**.
    * **Redirect URI (optional)** is filled in https://login.microsoftonline.com/common/oauth2/nativeclient.
    * Click **register**.
    * You will get a **Application (client) ID** in Overview, copy it.
* Click **Certificates & secrets**.
    * Click **+ New client secret**.
    * Fill in **Description** and select **Expires**(Futher is batter.In fact, we could select **never expire** before).
    * Click **Add**, and you will get a **Value**, copy it(If you close the web, you will never get it).
* Now Azure is ready.
### MS refresh_token
* Download the **GetToken.html** file in the project, or create a new HTML suffix document locally and copy the contents of **GetToken.html** into it. Then open it locally and get **refresh_token** according to the prompt inside.
### Server
* Upload the code named '**ApiOfRead.py**', '**ApiOfWrite.py**', '**UpdateToken.py**' to the **same directory** on the server.
* Download the file named '**ACCOUNT.txt**', fill the **Application (client) ID** into [""] after **client_id**, **Value** into [""] after **client_secret**, **refresh_token** into [""] after **ms_token**. Then upload it to the **same directory** on the server.

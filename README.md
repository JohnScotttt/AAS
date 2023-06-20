# AutoApiServer v1.0

<p style="font-size:60px;color:red">现可能失效，请注意</p>

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

## Usage
### MS E5
* I will not teach you how to get an MS E5 account, you need to get it by yourself online.
* If you have an MS E5 administrator account, next time we will use **it**, which ending with 'onmicrosoft.com'.
### Azure
* Open [Azure](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade), sign in the account.
* Click <kbd>+ New registration</kbd>.
    * Fill in **Name**, select the third item for the **supported account types**.
    * **Redirect URI (optional)** is filled in https://login.microsoftonline.com/common/oauth2/nativeclient.
    * Click <kbd>register</kbd>.
    * You will get a **Application (client) ID** in Overview, copy it.
* Click <kbd>Certificates & secrets</kbd>.
    * Click <kbd>+ New client secret</kbd>.
    * Fill in **Description** and select **Expires**(futher is better.In fact, I could select **never expire** before).
    * Click <kbd>Add</kbd>, and you will get a **Value**, copy it(If you close the web, you will never get it).
### MS refresh_token
* Download the ***GetToken.html*** file in the project, or create a new HTML suffix document locally and copy the contents of **GetToken.html** into it. Then open it locally and get **refresh_token** according to the prompt inside.
### Config
* ACCOUNT.txt
    * Download the file named ***ACCOUNT.txt***, fill the **Application (client) ID** into "" after **client_id**, fill the **Value** into "" after **client_secret**, fill the **refresh_token** into "" after **ms_token**. Then upload it to a directory on the server.
* EMAIL.txt
    * Download the file named ***EMAIL.txt***, fill your email into "" after **email**, fill your city(Capitalize initial) into "" after **city**. Then upload it to the **same directory** on the server.
### Server(If)
* .sh(If your server is ubuntu or other linux)
    * Download the files named ending with ***.sh***, fill your **directory absolute path** after **/** all of them. Then upload them to the **same directory** on the server.
* Python
    * Install **python** on your server or environment, usually python has already install on server.
    * You need to pip install requests, xlsxwriter(the most important. If one of something is missing, pip install it).
    * Upload the code named '***ApiOfRead.py***', '***ApiOfWrite.py***', '***UpdateToken.py***' to the **same directory** on the server.
* Cron
    * Type '**sudo crontab -e**' at the terminal, select your editor, type '**i**', then fill:
        ```shell
        10 10 * * 1,4,6 /(your directory absolute path)/UpdateToken.sh
        12 23 * * * /(your directory absolute path)/ApiOfWrite.sh
        12 */6 * * 1-5 /(your directory absolute path)/ApiOfRead.sh
        ```
    * Tap **ESC** on the keyboard, then type '**:wq**' to quit the editor, then tpye '**service cron restart**' to save cron.

## After that
* In theory, everything is OK if you follow the steps I give. In fect, you can type 'python **.py' to test them.
* Now it just only supports one account, one email and one city(because of my laziness). I don't guarantee that I will update them(maybe).
* Maybe it's too cumbersome to implement automation, but I'm too lazy to change them.

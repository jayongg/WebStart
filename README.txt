Quick web service:

1. Install nodejs, get it in your path.  From https://nodejs.org
2. run npm install
3. may need to run npm install -g gulp
4. To run the web server:
	gulp serve-web
	(server on port 8445, https)
	This will also compile and jsx (none being used now) and .ts files into .js, and copy to Stickers/dist/ folder.

	For insecure service, also run:
	gulp server-static-insecure
	(served on port 8444)
	This will serve the contents of serve-web onto 8444.

Make sure your firewall is taken care of if you are hitting this remotely.

for UWP:
Copy the manifests\stickers.xml to C:\Users\%USERNAME%\AppData\Local\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalState\AppData\Local\developer
Restart OneNote

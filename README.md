1. Install latest Google Chrome (option if it's installed. But if it's old, it should be updated)

2. Download ChromeDriver from https://googlechromelabs.github.io/chrome-for-testing/#stable in the project folder

    - MacOS:

        ```shell
        xattr -d com.apple.quarantine chromedriver
        ```

    - Linux:

        ```shell
        chmod +x chromedriver
        ```

3. Download google service account credential json file (credentials.json) in the project folder

    Refer to the following article:

    https://medium.com/@hitesh.thakur/how-to-upload-file-into-google-drive-via-python-using-service-account-a81f3ed54c66

4. Install Python 3.10 from https://www.python.org/downloads/release/python-3109/ (optional if you have already installed it)

5. Install necessary packages:

    ```shell
    pip install -r requirements.txt
    ```

6. Run the project:

    ```shell
    python main.py
    ```

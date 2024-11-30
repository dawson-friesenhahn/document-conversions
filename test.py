import requests
import os

os.environ["no_proxy"] = "127.0.0.1"

if __name__ == "__main__":
    files= {"file": open("TestPowerpoint.pptx", "rb")}

    resp= requests.get("http://127.0.0.1:5000")
    print(resp.text)

    resp= requests.post("http://127.0.0.1:5000/pptx-to-pdf", files=files)
    if not resp.ok:
        print(resp.text)
        print(resp.status_code)
    with open("output.pdf", "wb") as f:
        f.write(resp.content)


    print("done")
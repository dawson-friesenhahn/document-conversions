import requests
import os

os.environ["no_proxy"] = "127.0.0.1"

if __name__ == "__main__":
    files= {"file": ("Deleteme.pptx", open("DeleteMe.pptx", "rb"), 'multipart/form-data')}

    resp= requests.get("http://127.0.0.1:5000/")
    print(resp.text)

    data= {"garbage": "This is some garbage"}

    resp= requests.post("http://127.0.0.1:5000/pptx-to-pdf", data=data, files=files)


    print("done")
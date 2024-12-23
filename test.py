import requests
import os

os.environ["no_proxy"] = "127.0.0.1"

if __name__ == "__main__":
    # files = {"file": open("./test_files/TestPowerpoint.pptx", "rb")}
    # files= {"file": open("./test_files/NotARealPowerpoint.pptx", "rb")}
    # files= {"file": open("requirements.txt", "rb")}
    files = {"file": open("tlas.pptx", "rb")}

    resp = requests.post("http://127.0.0.1:5000/pptx-to-pdf", files=files)
    if not resp.ok:
        print(resp.text)
        print(resp.status_code)
    else:
        with open("output.pdf", "wb") as f:
            f.write(resp.content)
        print(f"saved file as {os.path.join(os.getcwd(), 'output.pdf')}")

    files = {"file": open("tlas.pptx", "rb")}
    resp = requests.post("http://127.0.0.1:5000/extract-pptx-text", files=files)
    if not resp.ok:
        print(resp.text)
        print(resp.status_code)
    else:
        print(f"Powerpoint contained text:\n{resp.text}")
    

    print("done")

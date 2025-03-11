import requests

url = "https://undermine.exchange/#eu-silvermoon/213469"  # Replace with the desired URL
response = requests.get(url)

if response.status_code == 200:
    html_content = response.text
    
    # Save HTML content to a file
    with open("webpage.html", "w", encoding="utf-8") as file:
        file.write(html_content)
    
    print("HTML content saved to webpage.html")
else:
    print(f"Failed to retrieve page, status code: {response.status_code}")

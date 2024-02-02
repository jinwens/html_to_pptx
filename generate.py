from jinja2 import Environment, FileSystemLoader
import json
import pdfkit 
import requests
from requests.auth import HTTPBasicAuth


def render_template(template_path, output_path, **kwargs):
    # Create a Jinja2 environment with the templates folder
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template(template_path)
    # Render the template with the provided variables
    output_content = template.render(**kwargs)
    # Write the rendered content to the output file
    with open(output_path, 'w') as output_file:
        output_file.write(output_content)

def convert_pdf_to_pptx(pdf_path, pptx_file, api_key):

    endpoint = "https://sandbox.zamzar.com/v1/jobs"
    #source_file = pdf_path
    target_format="pptx"
    files = {"source_file": open(pdf_path, "rb")}
    data = {
        "source_file_format": "pdf",
        "target_format": target_format
    }

    # Create a conversion job
    response = requests.post(endpoint, data=data, files=files, auth=HTTPBasicAuth(api_key, ''))
    #print(response.json())

    if response.status_code == 201:
        # Job created successfully
        job_id = response.json()["id"]
        print(f"Conversion job created with ID: {job_id}")

        # Wait for the job to finish
        job_status = "initialising"
        endpoint = "https://sandbox.zamzar.com/v1/jobs/{}".format(job_id)
        while job_status != "successful":
            response = requests.get(endpoint, auth=HTTPBasicAuth(api_key, ''))
            job_status = response.json()["status"]
            print(f"Job status: {job_status}")

        # Download the converted file
        file_id = response.json()["target_files"][0]['id']
        download_url = "https://sandbox.zamzar.com/v1/files/{}/content".format(file_id)   
        response = requests.get(download_url, stream=True, auth=HTTPBasicAuth(api_key, ''))

        # Save the converted file
        output_path = pptx_file
        with open(output_path, "wb") as output_file:
            output_file.write(response.content)

        print(f"Conversion complete. Output saved to: {output_path}")

    else:
        print(f"Failed to create conversion job. Status code: {response.status_code}")
        print(response.json())

if __name__ == "__main__":
    # Specify the template file and output file
    template_file = 'template.html'
    html_file = 'output.html'
    pdf_file = 'test.pdf'
    pptx_file = 'test.pptx'
    json_file = 'slide.json'
    # Read slide data
    with open(json_file, 'r') as json_file:
        data = json.load(json_file)
    # Provide template variables
    template_variables = {
        'title': data['head'],
        'userName': data['userName'],
        'images': data['images'][0]
    }
    # Generate HTML file
    render_template(template_file, html_file, **template_variables)
    # Generate pdf
    options = {
        "quiet": "",
        "javascript-delay": "5000",
        "page-width": "160mm",
        "page-height": "90.2mm",
        "dpi": 300,
        "margin-top": "0mm",
        "margin-right": "0mm",
        "margin-bottom": "0mm",
        "margin-left": "0mm",
        "no-outline": None,
        "encoding": "UTF-8",
        'enable-smart-shrinking': True,
        'enable-local-file-access': True
    }
    pdfkit.from_file(html_file,pdf_file,options=options)
    # Convert to pptx
    zamzar_api_key = '1a194542b8b1d930c5461f652e6ff55cc1f9591f'
    convert_pdf_to_pptx(pdf_file, pptx_file, zamzar_api_key)

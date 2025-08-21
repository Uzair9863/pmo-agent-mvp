# Week 2: Lambda handler for parsing RAID.xlsx from S3 and saving RAID.json
import json
import boto3
import openpyxl
import os
import tempfile

s3 = boto3.client("s3")

def lambda_handler(event, context):
    bucket_name = "pz-pmo-agent-data"   # your bucket name
    file_key = "RAID_fixed.xlsx"        # Excel file in S3
    output_key = "RAID.json"            # JSON output file in S3

    # Download Excel file from S3 to Lambda's /tmp directory
    tmp_path = os.path.join(tempfile.gettempdir(), "RAID_fixed.xlsx")
    s3.download_file(bucket_name, file_key, tmp_path)

    # Open Excel file with openpyxl
    wb = openpyxl.load_workbook(tmp_path)
    sheet = wb.active

    # Read all rows from the sheet
    rows = []
    for row in sheet.iter_rows(values_only=True):
        rows.append(list(row))

    # Convert to JSON
    json_data = {"rows": rows}

    # Save JSON back to S3
    s3.put_object(
        Bucket=bucket_name,
        Key=output_key,
        Body=json.dumps(json_data),
        ContentType="application/json"
    )

    return {
        "statusCode": 200,
        "body": json.dumps({
            "message": f"Parsed data saved to s3://{bucket_name}/{output_key}",
            "rows": rows
        })
    }


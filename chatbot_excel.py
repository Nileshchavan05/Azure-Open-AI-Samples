#pip install langchain openai pandas azure-storage-blob openpyxl

import pandas as pd
from azure.storage.blob import BlobServiceClient
from langchain.agents import create_pandas_dataframe_agent
from langchain.chat_models import ChatOpenAI
from langchain.chat_models import AzureChatOpenAI

# ---------- Azure Blob Setup ----------
AZURE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=your_account;AccountKey=your_key;EndpointSuffix=core.windows.net"
CONTAINER_NAME = "your-container"
BLOB_NAME = "monthly_sales.xlsx"
LOCAL_FILE_NAME = "sales_data.xlsx"

# Connect to Blob Storage and download the Excel file
blob_service_client = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
blob_client = blob_service_client.get_blob_client(container=CONTAINER_NAME, blob=BLOB_NAME)
with open(LOCAL_FILE_NAME, "wb") as file:
   file.write(blob_client.download_blob().readall())

# ---------- Read Excel into DataFrame ----------
df = pd.read_excel(LOCAL_FILE_NAME)

# ---------- Create LangChain Agent ----------
llm = ChatOpenAI(temperature=0, model_name="gpt-4")  # or use AzureChatOpenAI for Azure OpenAI
 OR
llm = AzureChatOpenAI(
   deployment_name="gpt-4-deployment",
   openai_api_version="2023-07-01-preview",
   openai_api_key="your-azure-api-key",
   openai_api_base="https://your-resource-name.openai.azure.com/",
)
agent = create_pandas_dataframe_agent(llm, df, verbose=True)

# ---------- Ask Questions ----------
question = "What is the total sales in the West region?"
response = agent.run(question)
print(response)


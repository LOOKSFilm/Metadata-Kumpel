from EditShareAPI import FlowMetadata, EsAuth
import json

EsAuth.login("10.0.77.13", "christoph", "looksfilm")

fields = FlowMetadata.getCustomMetadataFields().fields_dict

data = {
    "custom": {fields["000 Checked for Upload"]: "True"}
    }

print(FlowMetadata.updateAsset(5826455, json.dumps(data)))
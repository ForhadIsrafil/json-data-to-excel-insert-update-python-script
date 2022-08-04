import pandas as pd
import json
import glob

# TODO: READ JSON DATA START
json_files = glob.glob("jsons/*.json")
for path in json_files:
    json_file = open(path.replace("\\", "/"), "r")
    dict_data = json.loads(json_file.read())

    doc_ref = dict_data['extracted_metadata']["filename"].split(".pdf")[0]
    doc_title = dict_data['DocTitle'][0]
    rev = dict_data['Rev'][-1]
    purpose_of_issue = dict_data['PurposeOfIssue'][-1]
    status = dict_data.get('Status', "")

    # TODO: READ JSON DATA END

    # TODO: CHECK AND UPDATE DATA START
    df = pd.read_excel("doc_listing.xlsx")
    is_doc_exist = df.loc[(df["Doc Ref"] == doc_ref)].copy()
    # print(is_doc_exist.index[0])

    temp_dict_arr = []
    if len(is_doc_exist) == 0:
        # temp_dict_arr.append({"Doc Ref": doc_ref, "Doc Title": doc_title, "Rev": rev, "Purpose of Issue": purpose_of_issue,
        #                       "Status": status})
        temp_df = pd.DataFrame(
            [{"Doc Ref": doc_ref, "Doc Title": doc_title, "Rev": rev, "Purpose of Issue": purpose_of_issue,
              "Status": status[-1] if status != "" else ""}])
        adding_temp_df = pd.concat([df, temp_df], ignore_index=True)
        adding_temp_df.to_excel("doc_listing.xlsx", index=False)

    if len(is_doc_exist) > 0:
        if doc_title != "":
            df.at[is_doc_exist.index[0], "Doc Title"] = doc_title
        if rev != "":
            df.at[is_doc_exist.index[0], "Rev"] = rev
        if purpose_of_issue != "":
            df.at[is_doc_exist.index[0], "Purpose of Issue"] = purpose_of_issue
        if status != "":
            df.at[is_doc_exist.index[0], "Status"] = status[-1]

        df.to_excel("doc_listing.xlsx", index=False)
    # TODO: CHECK AND UPDATE DATA END

["Doc Ref", "Doc Title", "Rev", "Purpose of Issue", "Status"]

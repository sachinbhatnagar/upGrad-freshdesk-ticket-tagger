def extract_piped_data(input_str):
    each_line = input_str.replace("\n", " ").split("||")
    remap_and_extract = [msg.split("|") for msg in each_line]
    get_obj_from_array = []
    for msg_array in remap_and_extract:
        if len(msg_array) == 4:
            message, timecode, timestamp, user = msg_array

            message = message.strip()
            if str(message).startswith("="):
                message = str(message)[1:]

            get_obj_from_array.append(
                {
                    "message": message,
                    "timecode": timecode.strip(),
                    "timestamp": timestamp.strip(),
                    "user": user.strip(),
                }
            )
    return [i for i in get_obj_from_array if i is not None]

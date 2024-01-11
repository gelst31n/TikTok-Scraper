from tikapi import TikAPI, ValidationException, ResponseException
from datetime import datetime
import json
import pandas as pd
import statistics as stats

#API Key
api = TikAPI('ilQNohDEB0RJjqQrIgXXmCKzi20KAiI63ZYfoOAw0NbM4KY9')

#Sound ID from URL
sound_id = "7268424827940260651"

cursor = 0
df = pd.DataFrame(columns=['username', 'followers'])

#Get Videos Under Sound
sound = api.public.music(
    id= sound_id
)

sound_name = sound.json()["itemList"][0]["music"]["title"]

while(sound):
    sound_info = sound.json()
    results_json = {}

    #Get Individual Account Info
    for item in range(len(sound_info['itemList'])):

        username_exists = False

        if int(cursor) < 30:
            for key, value in results_json.items():
                if value["username"] == sound_info["itemList"][item]["author"]["uniqueId"]:
                    username_exists = True
                    break
        else:
            if sound_info["itemList"][item]["author"]["uniqueId"] in df['username'].values:
                username_exists = True

        if username_exists == False and 200 <= sound_info["itemList"][item]["authorStats"]["followerCount"] and sound_info["itemList"][item]["authorStats"]["followerCount"] <= 400000:

            results_json[item] = {"username": sound_info["itemList"][item]["author"]["uniqueId"]}
            results_json[item]["secUid"] =  sound_info["itemList"][item]["author"]["secUid"]
            results_json[item]["link"] = "https://tiktok.com/@" + results_json[item]["username"]
            results_json[item]["nickname"] =  sound_info["itemList"][item]["author"]["nickname"]
            results_json[item]["followers"] =  sound_info["itemList"][item]["authorStats"]["followerCount"]
            results_json[item]["posts under sound"] =  1

            print('Getting Information for ' + results_json[item]["username"])

            try:
                account = api.public.posts(
                    secUid = results_json[item]["secUid"]
                )
                account_info = account.json()

                results_json[item]["bio"] = account_info["itemList"][0]["author"]["signature"]

                views_list = []
                likes_list = []
                comments_list = []
                posts_in_past_month = 0

                for video in range(len(account_info["itemList"])):
                    createTime = account_info["itemList"][video]["createTime"]
                    time_since_post = datetime.today() - datetime.fromtimestamp(createTime)
                    if time_since_post.days <= 28:
                        views_list.append(account_info["itemList"][video]["stats"]["playCount"])
                        likes_list.append(account_info["itemList"][video]["stats"]["diggCount"])
                        comments_list.append(account_info["itemList"][video]["stats"]["commentCount"])
                        posts_in_past_month +=1

                if posts_in_past_month == 0:
                    results_json[item]["posts in last 28 days"] =  posts_in_past_month
                    results_json[item]["median views in last 28 days"] = 0
                    results_json[item]["mean views in last 28 days"] = 0
                    results_json[item]["median likes in last 28 days"] = 0
                    results_json[item]["median comments in last 28 days"] = 0
                    results_json[item]["median views per follower"] = 0
                else:
                    results_json[item]["posts in last 28 days"] =  posts_in_past_month
                    results_json[item]["median views in last 28 days"] = stats.median(views_list)
                    results_json[item]["mean views in last 28 days"] = stats.mean(views_list)
                    results_json[item]["median likes in last 28 days"] = stats.median(likes_list)
                    results_json[item]["median comments in last 28 days"] = stats.median(comments_list)
                    results_json[item]["median views per follower"] = results_json[item]["median views in last 28 days"]/results_json[item]["followers"]

                if "challenges" in account_info["itemList"][0]:
                    results_json[item]["tags"] = [account_info["itemList"][0]['challenges'][i]["title"] for i in range(len(account_info["itemList"][0]['challenges']))]
                else:
                    results_json[item]["tags"] = []

            except ResponseException as e:
                print('fail')

        elif username_exists == True:
            # For usernames found in results_json
            for key, value in results_json.items():
                if value["username"] == sound_info["itemList"][item]["author"]["uniqueId"]:
                    results_json[key]["posts under sound"] += 1
                    break  # exit the loop once the username is found and updated

            # For usernames found in DataFrame df
            if sound_info["itemList"][item]["author"]["uniqueId"] in df['username'].values:
                df.loc[df['username'] == sound_info["itemList"][item]["author"]["uniqueId"], 'posts under sound'] += 1

            print(sound_info["itemList"][item]["author"]["uniqueId"] + " Already In List")

        else:
            print(sound_info["itemList"][item]["author"]["uniqueId"] + " Doesn't Meet Follower Criteria")

    if cursor == 0:
        df = pd.DataFrame.from_dict(results_json, orient='index')
        df.to_csv(sound_name + '.csv')

    else:
        res = pd.DataFrame.from_dict(results_json, orient='index')
        df = df._append(res)
        df.to_csv(sound_name + '.csv')

    cursor = sound.json().get('cursor')
    print("Getting next items ", cursor)
    sound = sound.next_items()

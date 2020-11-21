import pandas as pd
import matplotlib.pyplot as plt

air_quality = pd.read_csv()
air_quality.rename(columns={"date.utc": "datetime"})
air_quality.head(n=10)

# column datetime as datetime objects instead of plain text
# pd.read_csv("../data/air_quality_no2_long.csv", parse_dates=["datetime"])
air_quality["datetime"] = pd.to_datetime(air_quality["datetime"])

air_quality["month"] = air_quality["datetime"].dt.month

air_quality.groupby([air_quality["datetime"].dt.weekday, "location"])["value"].mean()

fig, axs = plt.subplots(figsize=(12, 4))

air_quality.groupby(air_quality["datetime"].dt.hour)["value"].mean().plot(kind='bar', rot=0, ax=axs)

plt.xlabel("Hour of the day") 
plt.ylabel("$NO_2 (µg/m^3)$")

no_2 = air_quality.pivot(index="datetime", columns="location", values="value")
no_2.head()
no_2["2019-05-20":"2019-05-21"].plot()

monthly_max = no_2.resample("M").max()
monthly_max.index.freq
no_2.resample("D").mean().plot(style="-o", figsize=(10, 5))
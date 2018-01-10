import pandas as pd
import numpy as np
import xlrd
import glob
import csv

class Fish(object):
	def __init__(self, df):
		self.df = pd.DataFrame(data=df, copy=True)
		self.trialid = self.df.loc[df["Number of header lines:"] == "Trial name"]["35"].item().split()[1]
		self.arenaid = self.df.loc[self.df["Number of header lines:"] == "Arena name"]["35"].item().split()[1]
		self.shoalid = self.trialid + self.arenaid
		self.group = self.df.loc[self.df["Number of header lines:"] == "trt"]["35"].item()
		self.subjectid = self.df.loc[self.df["Number of header lines:"] == "Subject name"]["35"].item()
		self.headeridx = self.df[self.df["Number of header lines:"] == "Trial time"].index.item()
		self.parse(self.df)
	def parse(self, messy):
		messy.columns = messy.iloc[self.headeridx]
		messy.drop(messy.index[0:(self.headeridx+2)], inplace=True)
		exclude, keep = [], []
		for i in messy.columns.tolist():
			if "Recording time" in i:
				keep.append("Time")
			elif "Subject" in i and self.subjectid in i:
				exclude.append(i)
			elif "Subject" in i:
				keep.append(self.subjectid.split()[1] + "to" + i.split()[5])
			else:
				exclude.append(i)
		for i in exclude:
			messy.drop(i, axis=1, inplace=True)
		messy.columns = keep
		messy.reset_index(drop=True, inplace=True)
		messy.replace("-", np.NaN, inplace=True)
		messy["NND"] = messy[messy.columns[1:]].min(axis=1)
		messy.set_index(["Time"], drop=True, inplace=True)
		messy.drop(messy.columns[:3], axis=1, inplace=True)

def compile(shoal):
	sid = shoal[0].shoalid
	gr = shoal[0].group
	dfconv = []
	for s in shoal:
		dfconv.append(s.df)
	combo = pd.concat(dfconv, axis=1)
	combo["Avg"] = combo[combo.columns].mean(axis=1)
	combo.drop(combo.columns[:4], axis=1, inplace=True)
	combo.set_index(pd.to_datetime(combo.index * 1000, unit="ms"), inplace=True)
	avgmin = combo.resample("T").mean()
	minutes = []
	for m in range(len(avgmin)):
		minutes.append("Min" + str(m+1))
	avgmin["Minute"] = minutes
	avgmin["ShoalID"] = sid
	pivoted = avgmin.pivot(index="ShoalID", columns="Minute", values="Avg")
	pivoted["Mutant"] = gr
	return pivoted

def parse(imp):
	sheets = imp.sheet_names
	fish = []
	shoallist = []
	shoals = []
	out = []
	for sheet in sheets: 
		fish.append(Fish(imp.parse(sheet)))
	for f in fish:
		if f.shoalid not in shoallist:
			shoallist.append(f.shoalid)
			shoals.append([])
	for i,s in enumerate(shoallist): 
		for f in fish: 
			if f.shoalid == s: 
				shoals[i].append(f)
	for shoal in shoals:
		out.append(compile(shoal))
	return pd.concat(out, axis=0)

def main():
	summary = []
	files = glob.glob("*.xlsx")
	total = len(files)
	count = 0
	for file in files:
		print("Parsing " + file)
		try:
			summary.append(parse(pd.ExcelFile(file)))
			count += 1
			print(str(int((count/total)*100)) + "% done.")
		except:
			print("Failed.")
			continue
	final = pd.concat(summary, axis=0)
	return final

main().to_excel("result.xlsx")

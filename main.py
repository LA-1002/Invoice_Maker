import csv,operator

data = csv.reader(open('results_8-12-2022.csv'),delimiter=',')

data = sorted(data, key=operator.itemgetter(2))



brands = [];

for d in data:
    brands.append(d[1]);
    list(dict.fromkeys(brands))


for b in brands:
    




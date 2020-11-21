import sys
import numpy as np
import pandas


group_size = 7
size_sigma = 1
mix = ['gender', 'nationality']
match = ['age', 's1', 's2', 's3']
scalings = {'gender': 0.2, 'nationality': 0.2, 'age': 0.2, 's1': 0.2, 's2': 0.8, 's3': 1}

genders = {'Male 男': 1, 'Female 女': -1}


# PREPROCESS DATA
students_raw = pandas.read_excel('afternoon.xlsx', index_col=0)
students = students_raw.copy()

# clean gender
students.replace(to_replace=genders, inplace=True)
# clean s1
students.replace(to_replace=' / 100', value='', inplace=True, regex=True)
students['s1'] = students['s1'].astype(float)


# clean countries
def clean_country(country):
    if country is not np.nan:
        return country[:len(country) - 1 - country[::-1].index(' ')]
    else:
        return np.nan


students['nationality'] = students.nationality.apply(clean_country)

# code countries as eastern/western
with open('countries_east.txt') as f:
    asian_countries = f.readline().split(', ')
with open('countries_west.txt') as f:
    western_countries = f.readline().split(', ')

for ind in students.index:
    nationality = students['nationality'][ind]
    if nationality is not np.nan:
        if nationality in asian_countries:
            students.loc[ind, 'nationality'] = 1
        elif nationality in western_countries:
            students.loc[ind, 'nationality'] = -1
        else:
            sys.exit(f"{nationality} not in any list")

students['nationality'] = students['nationality'].astype(float)


# remove rows without valid scores
students.dropna(inplace=True)

# calculate z-scores
for match_key in match:
    students[match_key] = (students[match_key] - students[match_key].mean()) / students[match_key].std()

# re-scale
for key, scaling in scalings.items():
    students.loc[:, key] *= scaling


# SORT
num_groups = int(np.ceil(students.shape[0]/group_size))

# initialize groups

best_mean_distance = np.inf
best_grouping = None

for i in range(30):
    mean_distances = []
    students['group'] = int(-1)
    for group_num in range(num_groups):
        # select student far from mean as group seed
        means = students.loc[students.group == -1].mean()
        distances = ((students.loc[students.group == -1, match] - means).pow(2)).sum(axis=1)
        distances += ((students.loc[students.group == -1, mix] + means).pow(2).sum(axis=1))
        distances *= np.random.normal(loc=1, scale=0.5, size=len(distances))
        students.at[distances.idxmax(), 'group'] = group_num

        center = students.loc[distances.idxmax()]

        for member_num in range(1, group_size):
            if len(students.loc[students.group == -1]) == 0:
                continue
            # select new student with smallest distance
            distances = ((students.loc[students.group == -1, match] - center).pow(2)).sum(axis=1)
            distances += ((students.loc[students.group == -1, mix] + center).pow(2).sum(axis=1))
            students.at[distances.idxmin(), 'group'] = group_num

        means = students.loc[students.group == group_num].mean()
        distances = ((students.loc[students.group == group_num, match] - means).pow(2)).sum(axis=1)
        distances += ((students.loc[students.group == group_num, mix] + means).pow(2).sum(axis=1))
        mean_distances.append(distances.mean())

    new_mean_distance = np.mean(mean_distances)
    if new_mean_distance < best_mean_distance:
        best_mean_distance = new_mean_distance
        print(new_mean_distance)
        best_grouping = students['group']

students.loc[:, 'group'] = best_grouping


# calculate distances to group means and display them using asterisks
students_raw['d'] = np.nan
for group_num in range(num_groups):
    means = students.loc[students.group == group_num].mean()
    distances = ((students.loc[students.group == group_num, match] - means).pow(2)).sum(axis=1)
    distances += ((students.loc[students.group == group_num, mix] + means).pow(2).sum(axis=1))
    students_raw.loc[students.loc[students.group == group_num].index, 'd'] = distances

min_d = students_raw['d'].min()
max_d = students_raw['d'].max()


def format_d(d):
    if not pandas.isna(d):
        num_stars = int(round((d - min_d) / (max_d - min_d) * 5))
        return ''.join(['*' for _ in range(num_stars)])


students_raw['d'] = students_raw.d.apply(format_d)


# save and print groups sorted by mean s3
average_s3s = []
for group_num in range(num_groups):
    average_s3s.append(students.loc[students.group == group_num, 's3'].mean())
sorted_indices = np.argsort(average_s3s)

with pandas.ExcelWriter('sorted.xlsx', engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet('Groups')
    writer.sheets['Groups'] = worksheet

    for group_num in sorted_indices:
        print('\n')
        print(students_raw.loc[students.index[students.group == group_num]])
        students_raw.loc[students.index[students.group == group_num]].to_excel(writer, sheet_name='Groups',
                                                                               startrow=group_num * (group_size + 2))

    students_raw.loc[~ students_raw.index.isin(students.index)].to_excel(writer, sheet_name='Groups',
                                                                         startrow=num_groups * (group_size + 2) + 2)


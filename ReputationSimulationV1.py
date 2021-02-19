
import numpy as np
import pandas as pd
import collections
import openpyxl

details_file_name = 'ReputationDetails.xlsx'


congruency_df = pd.read_excel(details_file_name, 0, index_col=0)
congruency_df = congruency_df.transpose()
congruency_information = congruency_df.to_dict('list', into=collections.OrderedDict)
details_df = pd.read_excel(details_file_name, 1, index_col=0)
details_df = details_df.transpose()
details = details_df.to_dict('list', into=collections.OrderedDict)
initial_reputations_df = pd.read_excel(details_file_name, 2, index_col=0)
user_count = len(initial_reputations_df)
ecosystem_list = list(initial_reputations_df)


user_list = initial_reputations_df.index.tolist()
user_data_list = []
for user in user_list:
    user_data_list.append(user + '.xlsx')


for ecosystem_name in ecosystem_list:
    vars()[ecosystem_name] = pd.DataFrame(initial_reputations_df[ecosystem_name])
    eval(ecosystem_name).columns = ['T0']


def apply_weightage(ecosystem, time_period):
    ecosystem_index_value = ecosystem_list.index(ecosystem)
    current_ecosystem_df = eval(ecosystem)
    ecosystem_congruency = congruency_information[ecosystem]
    weighted_scores_dict = collections.OrderedDict()
    for user_name in user_list:
        user_name_index_value = user_list.index(user_name)
        calculation_df = pd.DataFrame()
        internal_reputation = current_ecosystem_df.at[user_name, time_period]
        calculation_df['InternalReputation'] = [internal_reputation]
        for external_ecosystem in ecosystem_list:
            working_ecosystem = eval(external_ecosystem)
            time_period_index_value = len(working_ecosystem.columns) - 1
            working_ecosystem_reputation = working_ecosystem.iloc[user_name_index_value, time_period_index_value]
            calculation_df[external_ecosystem] = [working_ecosystem_reputation]
        for congruency_value in ecosystem_congruency:
            score_location = ecosystem_congruency.index(congruency_value)
            score_name = ecosystem_list[score_location]
            if np.isnan(calculation_df.iloc[0, score_location + 1]):
                calculation_df[score_name] = calculation_df['InternalReputation'] * congruency_value
            else:
                calculation_df.iloc[0, score_location + 1] = calculation_df.iloc[0, score_location + 1] * \
                                                             congruency_value
        weighted_score = (calculation_df.sum(axis=1) - calculation_df['InternalReputation'])
        weighted_scores_dict[user_name] = weighted_score
    for key in weighted_scores_dict:
        current_ecosystem_df.at[key, time_period] = weighted_scores_dict[key][0]


def calculate_reputation(ecosystem, time_period):
    ecosystem_index_value = ecosystem_list.index(ecosystem)
    current_ecosystem_df = eval(ecosystem)
    ecosystem_details = details[ecosystem]
    scores_dict = collections.OrderedDict()
    for user_name in user_list:
        vote_data_file = str(user_name + '.xlsx')
        vote_df = pd.read_excel(vote_data_file, ecosystem_index_value, index_col=0)
        calculation_df = pd.DataFrame(vote_df[time_period])
        calculation_df.columns = ['Votes']
        calculation_df['UserReputation'] = current_ecosystem_df[current_ecosystem_df.columns[-1]]
        calculation_df['VoteWeight'] = calculation_df['UserReputation'].apply(lambda x: ecosystem_details[1] if
                                                                              x >= ecosystem_details[0]
                                                                              else (0.0 if np.isnan(x)
                                                                                    else x * ecosystem_details[2]))
        calculation_df['WeightedVotes'] = calculation_df['VoteWeight'] * calculation_df['Votes']
        possible_votes = calculation_df['VoteWeight'].sum()
        total_votes = calculation_df['WeightedVotes'].sum()
        calculated_votes = (total_votes / possible_votes) * 10
        scores_dict[user_name] = calculated_votes
    current_ecosystem_df[time_period] = np.nan
    for key in scores_dict:
        current_ecosystem_df.at[key, time_period] = scores_dict[key]
    apply_weightage(ecosystem, time_period)


def run_simulation():
    time_periods_to_simulate =int(input('How many time periods?'))
    for time in range(1, time_periods_to_simulate + 1):
        time_period = str('T' + str(time))
        for ecosystem in ecosystem_list:
            if time % details[ecosystem][4] == 0:
                calculate_reputation(ecosystem, time_period)
                ecosystem_to_print = eval(ecosystem)
                #print(ecosystem_to_print)
            else:
                pass


def export_sheets():
    with pd.ExcelWriter('export.xlsx') as writer:
        for ecosystem in ecosystem_list:
            export_sheet = eval(ecosystem)
            export_sheet.to_excel(writer, sheet_name=ecosystem)


run_simulation()
export_sheets()


















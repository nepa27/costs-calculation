from change_excel import delete_data_in_the_weekend, get_data_from_email, get_last_costs, write_data
from get_messages import get_messages_from_email


if __name__ == "__main__":
    new_values = get_data_from_email(get_messages_from_email())
    last_values = get_last_costs()

    write_data(new_values, last_values)
    delete_data_in_the_weekend(last_values)
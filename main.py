import change_excel
import get_messages


if __name__ == "__main__":
    change_excel.write_data(
        change_excel.get_data_from_email(
            get_messages.get_messages_from_email()),
        change_excel.get_last_costs())

    change_excel.delete_data_in_the_weekend(change_excel.get_last_costs())

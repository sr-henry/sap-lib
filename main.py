import sap

if __name__ == "__main__":
  # sap_connection_data = sap.attach()
    sap_connection_data = sap.create()

    if sap_connection_data:

        application, connection, session = sap_connection_data

        # script here

        sap.close(sap_connection_data)


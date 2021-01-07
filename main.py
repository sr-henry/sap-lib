import sap

if __name__ == "__main__":
    sap_connection_data = sap.attach()
    # sap_connection_data = sap.create()
    
    if sap_connection_data:   
        app, con, session = sap_connection_data

        # Script here
        session.StartTransaction("transaction_code")

        sap.close(sap_connection_data)
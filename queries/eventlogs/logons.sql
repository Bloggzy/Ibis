CREATE TABLE events AS
SELECT * FROM read_csv_auto('{{INPUT_CSV}}');

COPY (
    WITH base AS (
        SELECT
            TimeCreated,
            EventId,
            CASE
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 40 THEN 'RDPSessionDisconnectReason'
                WHEN Channel = 'Security' AND EventId = 4624 THEN 'LogonSuccess'
                WHEN Channel = 'Security' AND EventId = 4625 THEN 'LogonFailure'
                WHEN Channel = 'Security' AND EventId = 4634 THEN 'Logoff'
                WHEN Channel = 'Security' AND EventId = 4647 THEN 'UserInitiatedLogoff'
                WHEN Channel = 'Security' AND EventId = 4648 THEN 'ExplicitCredLogon'
                WHEN Channel = 'Security' AND EventId = 4672 THEN 'AdminPrivilegesAssigned'
                WHEN Channel = 'Security' AND EventId = 4776 THEN 'DCAuthenticationAttempt'
                WHEN Channel = 'Security' AND EventId = 4778 THEN 'RDPReconnect'
                WHEN Channel = 'Security' AND EventId = 4779 THEN 'RDPDisconnect'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' AND EventId = 1149 THEN 'RDPAuthSuccess'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 21 THEN 'RDPSessionLogon'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 22 THEN 'RDPShellStart'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 23 THEN 'RDPSessionLogoff'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 24 THEN 'RDPSessionDisconnect'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 25 THEN 'RDPSessionReconnect'
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 39 THEN 'RDPSessionDisconnectReason'
                WHEN Channel = 'System' AND EventId = 9009 THEN 'SystemLogoff'
                ELSE 'Other'
            END AS EventTypeLabel,
            RemoteHost,
            CASE
                WHEN Channel = 'Security' AND EventId IN (4624, 4625, 4634, 4647, 4648, 4672, 4776, 4778, 4779) AND PayloadData1 LIKE 'Target:%' THEN TRIM(REPLACE(PayloadData1, 'Target: ', ''))
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId IN (21, 23, 24, 25, 39) THEN UserName
                ELSE UserName
            END AS ExtractedUserName,
            TRY_CAST(TRIM(REPLACE(PayloadData2, 'LogonType ', '')) AS INTEGER) AS LogonTypeValue,
            CASE
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId IN (21, 23, 24, 25) THEN TRIM(REPLACE(PayloadData1, 'Session ID: ', ''))
                WHEN Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId = 39 THEN TRIM(REPLACE(PayloadData1, 'TargetSession: ', ''))
                ELSE NULL
            END AS SessionID,
            MapDescription,
            Channel,
            UserName,
            PayloadData1,
            PayloadData2,
            PayloadData3,
            PayloadData4,
            PayloadData5,
            PayloadData6,
            ExecutableInfo
        FROM events
        WHERE
            (Channel = 'Security' AND EventId IN (4624, 4625, 4634, 4647, 4648, 4672, 4776, 4778, 4779))
            OR (Channel = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' AND EventId = 1149)
            OR (Channel = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventId IN (21, 22, 23, 24, 25, 39, 40))
            OR (Channel = 'System' AND EventId = 9009)
    )
    SELECT
        TimeCreated,
        EventId,
        EventTypeLabel,
        RemoteHost,
        ExtractedUserName,
        CASE
            WHEN LogonTypeValue = 2 THEN '2 (Interactive)'
            WHEN LogonTypeValue = 3 THEN '3 (Network)'
            WHEN LogonTypeValue = 4 THEN '4 (Batch)'
            WHEN LogonTypeValue = 5 THEN '5 (Service)'
            WHEN LogonTypeValue = 7 THEN '7 (Unlock)'
            WHEN LogonTypeValue = 8 THEN '8 (NetworkCleartext)'
            WHEN LogonTypeValue = 9 THEN '9 (NewCredentials)'
            WHEN LogonTypeValue = 10 THEN '10 (RemoteInteractive)'
            WHEN LogonTypeValue = 11 THEN '11 (CachedInteractive)'
            ELSE CAST(LogonTypeValue AS TEXT)
        END AS LogonType,
        SessionID,
        MapDescription,
        Channel,
        UserName,
        PayloadData1,
        PayloadData2,
        PayloadData3,
        PayloadData4,
        PayloadData5,
        PayloadData6,
        ExecutableInfo
    FROM base
    ORDER BY TimeCreated ASC
) TO '{{OUTPUT_CSV}}' (HEADER, DELIMITER ',');

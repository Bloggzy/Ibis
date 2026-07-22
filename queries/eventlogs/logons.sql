CREATE TABLE events AS
SELECT * FROM read_csv_auto('{{INPUT_CSV}}', all_varchar = true);

COPY (
    WITH normalised AS (
        SELECT
            *,
            TRY_CAST(TRIM(EventId) AS INTEGER) AS EventIdValue,
            TRIM(Channel) AS ChannelValue
        FROM events
    ),
    base AS (
        SELECT
            TimeCreated,
            EventIdValue AS EventId,
            CASE
                WHEN ChannelValue = 'Security' AND EventIdValue = 4624 THEN 'LogonSuccess'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4625 THEN 'LogonFailure'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4634 THEN 'Logoff'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4647 THEN 'UserInitiatedLogoff'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4648 THEN 'ExplicitCredentialLogon'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4672 THEN 'AdminPrivilegesAssigned'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4776 THEN 'NTLMCredentialValidation'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4778 THEN 'WindowStationReconnect'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4779 THEN 'WindowStationDisconnect'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4800 THEN 'WorkstationLock'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4801 THEN 'WorkstationUnlock'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4802 THEN 'ScreenSaverInvoked'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4803 THEN 'ScreenSaverDismissed'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' AND EventIdValue = 1149 THEN 'RDPNetworkConnectionEstablished'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 21 THEN 'RDPSessionLogon'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 22 THEN 'RDPShellStart'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 23 THEN 'RDPSessionLogoff'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 24 THEN 'RDPSessionDisconnect'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 25 THEN 'RDPSessionReconnect'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 39 THEN 'SessionDisconnectedBySession'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 40 THEN 'SessionDisconnectReason'
                WHEN ChannelValue = 'System' AND Provider = 'User32' AND EventIdValue = 1074 THEN 'ShutdownOrRestartInitiated'
                WHEN ChannelValue = 'System' AND Provider = 'Microsoft-Windows-Kernel-General' AND EventIdValue = 12 THEN 'OperatingSystemStarted'
                WHEN ChannelValue = 'System' AND Provider = 'Microsoft-Windows-Kernel-General' AND EventIdValue = 13 THEN 'OperatingSystemShutdown'
                WHEN ChannelValue = 'System' AND Provider = 'Microsoft-Windows-Kernel-Power' AND EventIdValue = 41 THEN 'UnexpectedRestart'
                WHEN ChannelValue = 'System' AND Provider = 'EventLog' AND EventIdValue = 6005 THEN 'EventLogServiceStarted'
                WHEN ChannelValue = 'System' AND Provider = 'EventLog' AND EventIdValue = 6006 THEN 'EventLogServiceStopped'
                WHEN ChannelValue = 'System' AND Provider = 'EventLog' AND EventIdValue = 6008 THEN 'UnexpectedShutdown'
                WHEN ChannelValue IN ('Application', 'System') AND Provider = 'Desktop Window Manager' AND EventIdValue = 9009 THEN 'DesktopWindowManagerExit'
                ELSE 'Other'
            END AS EventTypeLabel,
            CASE
                WHEN ChannelValue = 'Security' AND EventIdValue = 4624 THEN 'SessionStart'
                WHEN ChannelValue = 'Security' AND EventIdValue IN (4625, 4648, 4776) THEN 'Authentication'
                WHEN ChannelValue = 'Security' AND EventIdValue IN (4634, 4647) THEN 'SessionEnd'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4672 THEN 'PrivilegeContext'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4778 THEN 'SessionResume'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4779 THEN 'SessionPause'
                WHEN ChannelValue = 'Security' AND EventIdValue IN (4800, 4801, 4802, 4803) THEN 'Presence'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' AND EventIdValue = 1149 THEN 'Authentication'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue IN (21, 22) THEN 'SessionStart'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 23 THEN 'SessionEnd'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue IN (24, 39, 40) THEN 'SessionPause'
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 25 THEN 'SessionResume'
                WHEN ChannelValue IN ('Application', 'System') THEN 'SystemBoundary'
                ELSE 'Other'
            END AS ActivityClass,
            CASE
                WHEN ChannelValue = 'Security' AND EventIdValue = 4625 THEN 'Failure'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4776
                    AND NULLIF(regexp_extract(Payload, '"@Name":"Status","#text":"([^"]+)"', 1), '') = '0x0'
                    THEN 'Success'
                WHEN ChannelValue = 'Security' AND EventIdValue = 4776 THEN 'Failure'
                WHEN ChannelValue = 'System' AND EventIdValue IN (41, 6008) THEN 'Unexpected'
                ELSE 'Observed'
            END AS Outcome,
            RemoteHost,
            CASE
                WHEN ChannelValue = 'Security'
                    AND EventIdValue IN (4624, 4625, 4634, 4648, 4776, 4800, 4801, 4802, 4803)
                    AND PayloadData1 LIKE 'Target:%'
                    THEN TRIM(REPLACE(PayloadData1, 'Target: ', ''))
                WHEN ChannelValue = 'Security' AND EventIdValue = 4647 AND UserName LIKE 'Target:%'
                    THEN TRIM(REPLACE(UserName, 'Target: ', ''))
                ELSE NULLIF(TRIM(UserName), '')
            END AS ExtractedUserName,
            CASE
                WHEN PayloadData2 LIKE 'LogonType %'
                    THEN TRY_CAST(TRIM(REPLACE(PayloadData2, 'LogonType ', '')) AS INTEGER)
                ELSE NULL
            END AS LogonTypeValue,
            COALESCE(
                NULLIF(
                    CASE
                        WHEN PayloadData3 LIKE 'LogonId:%' THEN TRIM(REPLACE(PayloadData3, 'LogonId:', ''))
                        WHEN ChannelValue = 'Security' AND EventIdValue = 4672 AND PayloadData2 LIKE 'LogonId:%'
                            THEN TRIM(REPLACE(PayloadData2, 'LogonId:', ''))
                        ELSE NULL
                    END,
                    ''
                ),
                NULLIF(regexp_extract(Payload, '"@Name":"TargetLogonId","#text":"([^"]+)"', 1), ''),
                NULLIF(regexp_extract(Payload, '"@Name":"LogonID","#text":"([^"]+)"', 1), ''),
                NULLIF(regexp_extract(Payload, '"@Name":"SubjectLogonId","#text":"([^"]+)"', 1), '')
            ) AS SecurityLogonID,
            CASE
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
                    AND EventIdValue IN (21, 23, 24, 25)
                    THEN TRIM(REPLACE(PayloadData1, 'Session ID: ', ''))
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 22
                    THEN json_extract_string(TRY_CAST(Payload AS JSON), '$.UserData.EventXML.SessionID')
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 39
                    THEN TRIM(REPLACE(PayloadData1, 'TargetSession: ', ''))
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 40
                    THEN TRIM(REPLACE(PayloadData1, 'Session: ', ''))
                WHEN ChannelValue = 'Security' AND EventIdValue IN (4800, 4801, 4802, 4803)
                    THEN NULLIF(regexp_extract(Payload, '"@Name":"SessionId","#text":"([^"]+)"', 1), '')
                ELSE NULL
            END AS TerminalSessionID,
            CASE
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 39
                    THEN TRIM(REPLACE(PayloadData2, 'Source: ', ''))
                ELSE NULL
            END AS RelatedSessionID,
            CASE
                WHEN ChannelValue = 'Security' AND EventIdValue IN (4778, 4779) THEN PayloadData1
                ELSE NULL
            END AS SessionName,
            CASE
                WHEN ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue = 40
                    THEN TRIM(REPLACE(PayloadData2, 'Reason: ', ''))
                ELSE NULL
            END AS DisconnectReason,
            RecordNumber,
            EventRecordId,
            Level,
            Provider,
            ChannelValue AS Channel,
            ProcessId,
            ThreadId,
            Computer,
            ChunkNumber,
            UserId,
            MapDescription,
            UserName,
            PayloadData1,
            PayloadData2,
            PayloadData3,
            PayloadData4,
            PayloadData5,
            PayloadData6,
            ExecutableInfo,
            HiddenRecord,
            SourceFile,
            Keywords,
            ExtraDataOffset,
            Payload
        FROM normalised
        WHERE
            (ChannelValue = 'Security' AND EventIdValue IN (4624, 4625, 4634, 4647, 4648, 4672, 4776, 4778, 4779, 4800, 4801, 4802, 4803))
            OR (ChannelValue = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' AND EventIdValue = 1149)
            OR (ChannelValue = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' AND EventIdValue IN (21, 22, 23, 24, 25, 39, 40))
            OR (ChannelValue = 'System' AND Provider = 'User32' AND EventIdValue = 1074)
            OR (ChannelValue = 'System' AND Provider = 'Microsoft-Windows-Kernel-General' AND EventIdValue IN (12, 13))
            OR (ChannelValue = 'System' AND Provider = 'Microsoft-Windows-Kernel-Power' AND EventIdValue = 41)
            OR (ChannelValue = 'System' AND Provider = 'EventLog' AND EventIdValue IN (6005, 6006, 6008))
            OR (ChannelValue IN ('Application', 'System') AND Provider = 'Desktop Window Manager' AND EventIdValue = 9009)
    )
    SELECT
        TimeCreated,
        EventId,
        EventTypeLabel,
        ActivityClass,
        Outcome,
        RemoteHost,
        ExtractedUserName,
        CASE
            WHEN LogonTypeValue = 0 THEN '0 (System)'
            WHEN LogonTypeValue = 2 THEN '2 (Interactive)'
            WHEN LogonTypeValue = 3 THEN '3 (Network)'
            WHEN LogonTypeValue = 4 THEN '4 (Batch)'
            WHEN LogonTypeValue = 5 THEN '5 (Service)'
            WHEN LogonTypeValue = 7 THEN '7 (Unlock)'
            WHEN LogonTypeValue = 8 THEN '8 (NetworkCleartext)'
            WHEN LogonTypeValue = 9 THEN '9 (NewCredentials)'
            WHEN LogonTypeValue = 10 THEN '10 (RemoteInteractive)'
            WHEN LogonTypeValue = 11 THEN '11 (CachedInteractive)'
            WHEN LogonTypeValue = 12 THEN '12 (CachedRemoteInteractive)'
            WHEN LogonTypeValue = 13 THEN '13 (CachedUnlock)'
            ELSE CAST(LogonTypeValue AS TEXT)
        END AS LogonType,
        SecurityLogonID,
        TerminalSessionID,
        RelatedSessionID,
        SessionName,
        DisconnectReason,
        RecordNumber,
        EventRecordId,
        Level,
        Provider,
        Channel,
        ProcessId,
        ThreadId,
        Computer,
        ChunkNumber,
        UserId,
        MapDescription,
        UserName,
        PayloadData1,
        PayloadData2,
        PayloadData3,
        PayloadData4,
        PayloadData5,
        PayloadData6,
        ExecutableInfo,
        HiddenRecord,
        SourceFile,
        Keywords,
        ExtraDataOffset,
        Payload
    FROM base
    ORDER BY TRY_CAST(TimeCreated AS TIMESTAMP) ASC, TRY_CAST(RecordNumber AS BIGINT) ASC
) TO '{{OUTPUT_CSV}}' (HEADER, DELIMITER ',');

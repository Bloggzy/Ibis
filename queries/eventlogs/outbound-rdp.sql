CREATE TABLE events AS
SELECT * FROM read_csv_auto('{{INPUT_CSV}}');

COPY (
    SELECT
        TimeCreated,
        Computer,
        EventId,
        MapDescription,
        CASE
            WHEN EventId = 1102 THEN regexp_extract(PayloadData1, 'Address: ([^ ]+)', 1)
            WHEN EventId = 1024 THEN regexp_extract(PayloadData1, 'Dest: ([^ ]+)', 1)
            ELSE NULL
        END AS Address,
        CASE
            WHEN EventId = 1029 THEN regexp_extract(PayloadData1, 'Target \(encoded\): (.+)', 1)
            ELSE NULL
        END AS EncodedUsername,
        CASE
            WHEN EventId = 1027 THEN regexp_extract(PayloadData1, 'Domain: ([^ ]+)', 1)
            ELSE NULL
        END AS Domain,
        CASE
            WHEN EventId = 1027 THEN regexp_extract(PayloadData2, 'Session ID: (\d+)', 1)
            ELSE NULL
        END AS SessionID,
        UserId,
        PayloadData1,
        PayloadData2,
        PayloadData3,
        PayloadData4,
        PayloadData5,
        PayloadData6,
        ExecutableInfo,
        Channel,
        SourceFile
    FROM events
    WHERE Channel = 'Microsoft-Windows-TerminalServices-RDPClient/Operational'
        AND EventId IN (1024, 1025, 1026, 1027, 1029, 1101, 1102, 1103, 1104, 1105)
    ORDER BY TimeCreated ASC
) TO '{{OUTPUT_CSV}}' (HEADER, DELIMITER ',');

CREATE TABLE events AS
SELECT * FROM read_csv_auto('{{INPUT_CSV}}');

COPY (
    SELECT
        Channel,
        MIN(TimeCreated) AS OldestEvent,
        MAX(TimeCreated) AS NewestEvent
    FROM events
    WHERE Channel IN (
        'Application',
        'Microsoft-Windows-PowerShell/Operational',
        'Microsoft-Windows-NetworkProfile/Operational',
        'Microsoft-Windows-TaskScheduler/Operational',
        'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational',
        'Microsoft-Windows-TerminalServices-RDPClient/Operational',
        'Microsoft-Windows-Windows Defender/Operational',
        'Microsoft-Windows-WMI-Activity/Operational',
        'Windows PowerShell',
        'Security',
        'System'
    )
    GROUP BY Channel
    ORDER BY Channel
) TO '{{OUTPUT_CSV}}' (HEADER, DELIMITER ',');

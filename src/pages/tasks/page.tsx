"use client"

// Enhanced TeamsMeeting.tsx with teams.live.com URL targeting and UTC timezone selection
import { useState, useEffect } from "react";
import { AuthenticatedTemplate } from "@azure/msal-react";
import { 
  createValidTeamsMeeting, 
  createTeamsMeetingWithSpecificFormat,
  createPersonalTeamsMeeting,
  detectAccountType 
} from "../../GraphService";
import { useAppContext } from "../../AppContext";

// Common timezones with UTC offsets
const TIMEZONES = [
  { value: 'UTC', label: 'UTC (Coordinated Universal Time)', offset: '+00:00' },
  { value: 'America/New_York', label: 'Eastern Time (ET)', offset: '-05:00/-04:00' },
  { value: 'America/Chicago', label: 'Central Time (CT)', offset: '-06:00/-05:00' },
  { value: 'America/Denver', label: 'Mountain Time (MT)', offset: '-07:00/-06:00' },
  { value: 'America/Los_Angeles', label: 'Pacific Time (PT)', offset: '-08:00/-07:00' },
  { value: 'Europe/London', label: 'GMT/BST (London)', offset: '+00:00/+01:00' },
  { value: 'Europe/Paris', label: 'CET/CEST (Paris)', offset: '+01:00/+02:00' },
  { value: 'Europe/Berlin', label: 'CET/CEST (Berlin)', offset: '+01:00/+02:00' },
  { value: 'Asia/Tokyo', label: 'JST (Tokyo)', offset: '+09:00' },
  { value: 'Asia/Shanghai', label: 'CST (Shanghai)', offset: '+08:00' },
  { value: 'Asia/Kolkata', label: 'IST (India)', offset: '+05:30' },
  { value: 'Australia/Sydney', label: 'AEDT/AEST (Sydney)', offset: '+11:00/+10:00' },
  { value: 'Pacific/Auckland', label: 'NZDT/NZST (Auckland)', offset: '+13:00/+12:00' },
];

export default function EnhancedTeamsMeeting() {
  const app = useAppContext();
  const [subject, setSubject] = useState("");
  const [start, setStart] = useState("");
  const [end, setEnd] = useState("");
  const [timezone, setTimezone] = useState("UTC");
  const [attendees, setAttendees] = useState("");
  const [body, setBody] = useState("");
  const [meetingDetails, setMeetingDetails] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [permissionError, setPermissionError] = useState<string | null>(null);
  const [creationMethod, setCreationMethod] = useState<'enhanced' | 'specific_format' | 'personal'>('specific_format');
  const [accountInfo, setAccountInfo] = useState<any>(null);

  // Detect user's timezone on component mount
  useEffect(() => {
    const userTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const matchingTimezone = TIMEZONES.find(tz => tz.value === userTimezone);
    if (matchingTimezone) {
      setTimezone(userTimezone);
    }
  }, []);

  // Check account type when component loads
  const checkAccountType = async () => {
    if (app.authProvider) {
      try {
        const info = await detectAccountType(app.authProvider);
        setAccountInfo(info);
      } catch (error) {
        console.error("Error detecting account type:", error);
      }
    }
  };

  // Run account check on mount
  useEffect(() => {
    checkAccountType();
  }, [app.authProvider]);

  // Helper function to convert local datetime to specified timezone
  const convertToTimezone = (datetimeLocal: string, targetTimezone: string): Date => {
    const localDate = new Date(datetimeLocal);
    
    if (targetTimezone === 'UTC') {
      return new Date(localDate.getTime() - (localDate.getTimezoneOffset() * 60000));
    }
    
    // For other timezones, we'll create the date in the target timezone
    // This is a simplified approach - for production, consider using a proper date library like date-fns-tz
    return localDate;
  };

  // Helper function to get current datetime in local format for input
  const getCurrentDateTime = () => {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    return now.toISOString().slice(0, 16);
  };

  // Set default times when component loads
  useEffect(() => {
    if (!start) {
      const now = new Date();
      now.setHours(now.getHours() + 1, 0, 0, 0); // Next hour
      setStart(getCurrentDateTime());
    }
    if (!end) {
      const now = new Date();
      now.setHours(now.getHours() + 2, 0, 0, 0); // Hour after start
      const endTime = new Date(now);
      endTime.setMinutes(endTime.getMinutes() - endTime.getTimezoneOffset());
      setEnd(endTime.toISOString().slice(0, 16));
    }
  }, []);

  const handleCreateMeeting = async () => {
    setIsLoading(true);
    setPermissionError(null);
    
    try {
      // Convert datetime-local strings to Date objects in the specified timezone
      const startDate = convertToTimezone(start, timezone);
      const endDate = convertToTimezone(end, timezone);
      
      // Validate dates
      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        throw new Error("Please provide valid start and end dates");
      }
      
      if (startDate >= endDate) {
        throw new Error("Start date must be before end date");
      }

      if (!subject.trim()) {
        throw new Error("Meeting subject is required");
      }

      // Parse attendee emails
      const attendeeEmails = attendees
        .split(',')
        .map(email => email.trim())
        .filter(email => email.length > 0);

      let meeting;
      
      switch (creationMethod) {
        case 'personal':
          console.log("Using personal Teams meeting creation (targeting teams.live.com)...");
          meeting = await createPersonalTeamsMeeting(
            app.authProvider!,
            subject,
            startDate,
            endDate,
            attendeeEmails,
            body,
            timezone
          );
          break;
          
        case 'specific_format':
          console.log("Using format-specific Teams meeting creation...");
          meeting = await createTeamsMeetingWithSpecificFormat(
            app.authProvider!,
            subject,
            startDate,
            endDate,
            attendeeEmails,
            body,
            timezone
          );
          break;
          
        case 'enhanced':
        default:
          console.log("Using enhanced Teams meeting creation with fixed API calls...");
          meeting = await createTeamsMeetingWithSpecificFormat(
            app.authProvider!,
            subject,
            startDate,
            endDate,
            attendeeEmails,
            body,
            timezone
          );
          break;
      }
      
      setMeetingDetails(meeting);
      console.log("Meeting created successfully:", meeting);
      
      // Re-check account info after successful creation
      checkAccountType();
      
    } catch (error: any) {
      console.error("Error creating Teams meeting:", error);
      
      // Check if this is a permission-related error
      if (error.message?.toLowerCase().includes('permission') || 
          error.message?.toLowerCase().includes('insufficient privileges') ||
          error.message?.toLowerCase().includes('unauthorized') ||
          error.code === 'Forbidden' ||
          error.statusCode === 403) {
        setPermissionError("Insufficient permissions to create Teams meetings. Please contact your administrator to grant the required permissions.");
      } else {
        app.displayError!(error.message);
      }
    } finally {
      setIsLoading(false);
    }
  };

  const resetForm = () => {
    setSubject("");
    setStart("");
    setEnd("");
    setTimezone("UTC");
    setAttendees("");
    setBody("");
    setMeetingDetails(null);
    setPermissionError(null);
  };

  const copyToClipboard = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      // You could add a toast notification here
    } catch (err) {
      console.error('Failed to copy text: ', err);
    }
  };

  // Format timezone display
  const formatTimezoneDisplay = (timezoneValue: string) => {
    const tz = TIMEZONES.find(t => t.value === timezoneValue);
    return tz ? `${tz.label} (${tz.offset})` : timezoneValue;
  };

  return (
    <AuthenticatedTemplate>
      <div className="p-6 space-y-4 max-w-4xl mx-auto">
        <div className="flex items-center justify-between">
          <h2 className="text-2xl font-bold text-gray-800">Create Teams Meeting</h2>
          
          {/* Account Type Display */}
          {accountInfo && (
            <div className="text-sm text-gray-600 bg-gray-100 px-3 py-2 rounded-lg">
              Account: {accountInfo.isPersonalAccount ? 'Personal' : 'Business'} 
              <span className="ml-2 text-xs">
                (Expected: {accountInfo.expectedUrlFormat})
              </span>
            </div>
          )}
        </div>
        
        {/* URL Format Information */}
        <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
          <div className="flex items-center mb-2">
            <h3 className="font-medium text-blue-800">Teams Meeting URL Formats</h3>
          </div>
          <div className="text-sm text-blue-700 space-y-2">
            <p>
              <span className="font-medium">teams.live.com</span> - Generated for personal Microsoft accounts (outlook.com, hotmail.com, live.com)
            </p>
            <p>
              <span className="font-medium">teams.microsoft.com</span> - Generated for business/organizational accounts
            </p>
            <p className="text-xs text-blue-600 mt-2">
              The URL format is determined by your account type. Use "Personal Teams Meeting" method below to specifically target teams.live.com URLs.
            </p>
          </div>
        </div>

        {/* Creation Method Selection */}
        <div className="space-y-3">
          <label className="block text-sm font-medium text-gray-700">Creation Method:</label>
          <div className="space-y-2">
            <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-gray-50">
              <input
                type="radio"
                name="creationMethod"
                value="specific_format"
                checked={creationMethod === 'specific_format'}
                onChange={(e) => setCreationMethod(e.target.value as any)}
                className="w-4 h-4 text-blue-600"
              />
              <div>
                <div className="font-medium text-gray-800">Format-Specific Creation</div>
                <div className="text-sm text-gray-600">
                  Detects account type and uses appropriate method for desired URL format
                </div>
              </div>
            </label>

            <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-gray-50">
              <input
                type="radio"
                name="creationMethod"
                value="personal"
                checked={creationMethod === 'personal'}
                onChange={(e) => setCreationMethod(e.target.value as any)}
                className="w-4 h-4 text-blue-600"
              />
              <div>
                <div className="font-medium text-gray-800">Personal Teams Meeting</div>
                <div className="text-sm text-gray-600">
                  Specifically targets teams.live.com URL format (requires personal account)
                </div>
              </div>
            </label>

            <label className="flex items-center space-x-3 p-3 border rounded-lg cursor-pointer hover:bg-gray-50">
              <input
                type="radio"
                name="creationMethod"
                value="enhanced"
                checked={creationMethod === 'enhanced'}
                onChange={(e) => setCreationMethod(e.target.value as any)}
                className="w-4 h-4 text-blue-600"
              />
              <div>
                <div className="font-medium text-gray-800">Enhanced Standard Method</div>
                <div className="text-sm text-gray-600">
                  Multiple fallback methods for reliable meeting creation
                </div>
              </div>
            </label>
          </div>
        </div>
        
        <div className="space-y-4">
          {/* Subject */}
          <div>
            <label className="block text-sm font-medium mb-1 text-gray-700">
              Meeting Subject *
            </label>
            <input
              type="text"
              placeholder="Enter meeting subject"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              required
            />
          </div>

          {/* Timezone Selection */}
          <div>
            <label className="block text-sm font-medium mb-1 text-gray-700">
              Timezone *
            </label>
            <select
              value={timezone}
              onChange={(e) => setTimezone(e.target.value)}
              className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            >
              {TIMEZONES.map((tz) => (
                <option key={tz.value} value={tz.value}>
                  {tz.label} ({tz.offset})
                </option>
              ))}
            </select>
            <p className="text-xs text-gray-500 mt-1">
              Selected timezone: {formatTimezoneDisplay(timezone)}
            </p>
          </div>

          {/* Date/Time Row */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium mb-1 text-gray-700">
                Start Time * (in {timezone})
              </label>
              <input
                type="datetime-local"
                value={start}
                onChange={(e) => setStart(e.target.value)}
                className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                required
              />
              {start && (
                <p className="text-xs text-gray-500 mt-1">
                  UTC: {new Date(start).toISOString().replace('T', ' ').substring(0, 19)}
                </p>
              )}
            </div>

            <div>
              <label className="block text-sm font-medium mb-1 text-gray-700">
                End Time * (in {timezone})
              </label>
              <input
                type="datetime-local"
                value={end}
                onChange={(e) => setEnd(e.target.value)}
                className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                required
              />
              {end && (
                <p className="text-xs text-gray-500 mt-1">
                  UTC: {new Date(end).toISOString().replace('T', ' ').substring(0, 19)}
                </p>
              )}
            </div>
          </div>

          {/* Quick Time Buttons */}
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => {
                const now = new Date();
                now.setHours(now.getHours() + 1, 0, 0, 0);
                const startTime = new Date(now);
                const endTime = new Date(now);
                endTime.setHours(endTime.getHours() + 1);
                
                startTime.setMinutes(startTime.getMinutes() - startTime.getTimezoneOffset());
                endTime.setMinutes(endTime.getMinutes() - endTime.getTimezoneOffset());
                
                setStart(startTime.toISOString().slice(0, 16));
                setEnd(endTime.toISOString().slice(0, 16));
              }}
              className="px-3 py-1 text-sm bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
            >
              Next Hour (1h)
            </button>
            <button
              type="button"
              onClick={() => {
                const now = new Date();
                now.setDate(now.getDate() + 1);
                now.setHours(9, 0, 0, 0); // 9 AM tomorrow
                const startTime = new Date(now);
                const endTime = new Date(now);
                endTime.setHours(10, 0, 0, 0); // 10 AM tomorrow
                
                startTime.setMinutes(startTime.getMinutes() - startTime.getTimezoneOffset());
                endTime.setMinutes(endTime.getMinutes() - endTime.getTimezoneOffset());
                
                setStart(startTime.toISOString().slice(0, 16));
                setEnd(endTime.toISOString().slice(0, 16));
              }}
              className="px-3 py-1 text-sm bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
            >
              Tomorrow 9 AM (1h)
            </button>
          </div>

          {/* Attendees */}
          <div>
            <label className="block text-sm font-medium mb-1 text-gray-700">
              Attendee Emails
            </label>
            <input
              type="text"
              placeholder="Enter email addresses separated by commas (e.g., john@example.com, jane@example.com)"
              value={attendees}
              onChange={(e) => setAttendees(e.target.value)}
              className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            />
            <p className="text-xs text-gray-500 mt-1">
              Separate multiple emails with commas
            </p>
          </div>

          {/* Meeting Description */}
          <div>
            <label className="block text-sm font-medium mb-1 text-gray-700">
              Meeting Description
            </label>
            <textarea
              placeholder="Enter meeting description or agenda"
              value={body}
              onChange={(e) => setBody(e.target.value)}
              rows={4}
              className="border border-gray-300 p-3 rounded-lg w-full focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-vertical"
            />
          </div>

          {/* Action Buttons */}
          <div className="flex gap-3">
            <button
              onClick={handleCreateMeeting}
              disabled={!subject || !start || !end || isLoading}
              className="flex-1 bg-blue-600 text-white px-6 py-3 rounded-lg font-medium hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
            >
              {isLoading ? "Creating Teams Meeting..." : "Create Teams Meeting"}
            </button>
            
            {meetingDetails && (
              <button
                onClick={resetForm}
                className="px-6 py-3 bg-gray-200 text-gray-700 rounded-lg font-medium hover:bg-gray-300 transition-colors"
              >
                New Meeting
              </button>
            )}
          </div>
        </div>

        {/* Permission Warning */}
        {permissionError && (
          <div className="mt-4 p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
            <div className="flex items-start">
              <div className="w-5 h-5 text-yellow-600 mr-2 mt-0.5">
                <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.732-.833-2.5 0L4.268 15.5c-.77.833.192 2.5 1.732 2.5z" />
                </svg>
              </div>
              <div>
                <h4 className="font-medium text-yellow-800">Permission Notice</h4>
                <p className="text-sm text-yellow-700 mt-1">{permissionError}</p>
                <div className="mt-2 text-xs text-yellow-600">
                  <p><strong>Required permissions for Teams meetings:</strong></p>
                  <ul className="list-disc list-inside mt-1">
                    <li>OnlineMeetings.ReadWrite (essential for join URLs)</li>
                    <li>Calendars.ReadWrite</li>
                    <li>User.Read</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Success Message with URL Format Information */}
        {meetingDetails && meetingDetails.success && (
          <div className="mt-6 p-6 bg-green-50 border border-green-200 rounded-lg">
            <div className="flex items-center mb-4">
              <div className="w-10 h-10 bg-green-100 rounded-full flex items-center justify-center mr-3">
                <svg className="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <div>
                <h3 className="font-bold text-green-800">Teams Meeting Created Successfully!</h3>
                <p className="text-sm text-green-600">
                  Created using: {meetingDetails.method?.replace(/_/g, ' ').toUpperCase()}
                  {meetingDetails.retriesNeeded && ` (${meetingDetails.retriesNeeded} retries needed)`}
                </p>
                {meetingDetails.urlFormat && (
                  <p className="text-sm text-green-600">
                    URL Format: <span className="font-mono font-bold">{meetingDetails.urlFormat}</span>
                    {meetingDetails.isDesiredFormat === true && (
                      <span className="ml-2 text-green-700 font-bold">✓ teams.live.com achieved!</span>
                    )}
                    {meetingDetails.isDesiredFormat === false && (
                      <span className="ml-2 text-orange-600">Different format generated</span>
                    )}
                  </p>
                )}
              </div>
            </div>
            
            <div className="space-y-4">
              {/* Meeting Details */}
              <div className="bg-white p-4 rounded-lg border border-green-200">
                <p className="font-medium text-green-800 mb-2">Meeting Details:</p>
                <p className="text-green-700 font-semibold">{meetingDetails.subject}</p>
                <p className="text-sm text-gray-600">
                  {new Date(meetingDetails.startTime.dateTime).toLocaleString()} - 
                  {new Date(meetingDetails.endTime.dateTime).toLocaleString()}
                </p>
                <p className="text-sm text-gray-600">
                  Timezone: {meetingDetails.startTime.timeZone || timezone}
                </p>
                {meetingDetails.accountType && (
                  <p className="text-sm text-gray-600 mt-1">
                    Account Type: {meetingDetails.accountType.isPersonalAccount ? 'Personal' : 'Business'} 
                    ({meetingDetails.accountType.userPrincipalName})
                  </p>
                )}
              </div>
              
              {/* Join URL - Most Important with Format Highlighting */}
              <div className="bg-white p-4 rounded-lg border border-green-200">
                <div className="flex items-center justify-between mb-2">
                  <p className="font-medium text-green-800">Teams Meeting Join URL:</p>
                  {meetingDetails.joinWebUrl?.includes('teams.live.com') && (
                    <span className="bg-purple-100 text-purple-800 text-xs px-2 py-1 rounded-full font-medium">
                      teams.live.com ✓
                    </span>
                  )}
                  {meetingDetails.joinWebUrl?.includes('teams.microsoft.com') && (
                    <span className="bg-blue-100 text-blue-800 text-xs px-2 py-1 rounded-full font-medium">
                      teams.microsoft.com
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2 mb-2">
                  <a
                    href={meetingDetails.joinWebUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex-1 bg-blue-600 text-white text-center py-3 px-4 rounded hover:bg-blue-700 transition-colors font-medium"
                  >
                    Join Teams Meeting Now
                  </a>
                  <button
                    onClick={() => copyToClipboard(meetingDetails.joinWebUrl)}
                    className="px-3 py-3 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
                    title="Copy join URL"
                  >
                    Copy URL
                  </button>
                </div>
                <p className="text-xs text-gray-500 break-all font-mono bg-gray-50 p-2 rounded">
                  {meetingDetails.joinWebUrl}
                </p>
              </div>

              {/* Meeting ID */}
              {meetingDetails.meetingId && (
                <div className="bg-white p-4 rounded-lg border border-green-200">
                  <p className="font-medium text-green-800 mb-2">Meeting ID:</p>
                  <div className="flex items-center gap-2">
                    <p className="text-green-700 font-mono text-lg flex-1">{meetingDetails.meetingId}</p>
                    <button
                      onClick={() => copyToClipboard(meetingDetails.meetingId)}
                      className="px-3 py-1 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors text-sm"
                    >
                      Copy
                    </button>
                  </div>
                </div>
              )}

              {/* Conference Details */}
              {(meetingDetails.tollNumber || meetingDetails.tollFreeNumber || meetingDetails.conferenceId) && (
                <div className="bg-white p-4 rounded-lg border border-green-200">
                  <p className="font-medium text-green-800 mb-2">Phone Conference Details:</p>
                  {meetingDetails.tollFreeNumber && (
                    <p className="text-green-700">Toll-free: {meetingDetails.tollFreeNumber}</p>
                  )}
                  {meetingDetails.tollNumber && (
                    <p className="text-green-700">Toll: {meetingDetails.tollNumber}</p>
                  )}
                  {meetingDetails.conferenceId && (
                    <p className="text-green-700">Conference ID: {meetingDetails.conferenceId}</p>
                  )}
                  {meetingDetails.dialInUrl && (
                    <p className="text-green-700">
                      <a href={meetingDetails.dialInUrl} target="_blank" rel="noopener noreferrer" className="text-blue-600 underline">
                        View full dial-in details
                      </a>
                    </p>
                  )}
                </div>
              )}

              {/* Attendees List */}
              {meetingDetails.attendees && meetingDetails.attendees.length > 0 && (
                <div className="bg-white p-4 rounded-lg border border-green-200">
                  <p className="font-medium text-green-800 mb-2">Invited Attendees:</p>
                  <ul className="space-y-1">
                    {meetingDetails.attendees.map((attendee: any, index: number) => (
                      <li key={index} className="text-green-700 text-sm">
                        {attendee.emailAddress?.name || attendee.emailAddress?.address}
                        <span className="text-gray-500 ml-1">({attendee.emailAddress?.address})</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {/* Technical Details */}
              <div className="bg-gray-50 p-4 rounded-lg border">
                <p className="font-medium text-gray-800 mb-2">Technical Details:</p>
                <div className="text-sm text-gray-600 space-y-1">
                  <p><span className="font-medium">Meeting ID:</span> {meetingDetails.id}</p>
                  <p><span className="font-medium">Calendar Event ID:</span> {meetingDetails.eventId}</p>
                  <p><span className="font-medium">Creation Method:</span> {meetingDetails.method}</p>
                  <p><span className="font-medium">URL Format:</span> {meetingDetails.urlFormat || 'Unknown'}</p>
                  <p><span className="font-medium">Timezone:</span> {meetingDetails.startTime?.timeZone || timezone}</p>
                  {meetingDetails.retriesNeeded && (
                    <p><span className="font-medium">Retries Needed:</span> {meetingDetails.retriesNeeded}</p>
                  )}
                </div>
              </div>
            </div>

            {/* URL Format Explanation */}
            <div className="mt-4 p-4 bg-purple-50 rounded-lg border-l-4 border-purple-400">
              <div className="flex items-start">
                <div className="mr-2">📋</div>
                <div>
                  <p className="text-sm text-purple-800 font-medium">About URL Formats:</p>
                  <ul className="text-xs text-purple-700 mt-1 space-y-1">
                    <li>• <strong>teams.live.com</strong> URLs are generated for personal Microsoft accounts</li>
                    <li>• <strong>teams.microsoft.com</strong> URLs are generated for business accounts</li>
                    <li>• Both formats provide the same Teams meeting functionality</li>
                    <li>• The format is determined by your Microsoft account type, not the creation method</li>
                    <li>• To get teams.live.com URLs, you need to sign in with a personal Microsoft account (@outlook.com, @hotmail.com, @live.com)</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Loading State */}
        {isLoading && (
          <div className="mt-6 p-6 bg-blue-50 border border-blue-200 rounded-lg">
            <div className="flex items-center">
              <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600 mr-3"></div>
              <div>
                <p className="text-blue-800 font-medium">Creating Teams Meeting...</p>
                <p className="text-sm text-blue-600">
                  {creationMethod === 'personal' && "Targeting teams.live.com URL format..."}
                  {creationMethod === 'specific_format' && "Using format-specific creation method..."}
                  {creationMethod === 'enhanced' && "Using enhanced creation with multiple fallbacks..."}
                </p>
                <p className="text-xs text-blue-600 mt-1">
                  Creating meeting in {timezone} timezone
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Troubleshooting Tips */}
        <div className="mt-6 p-4 bg-gray-50 rounded-lg border">
          <h4 className="font-medium text-gray-800 mb-2">Getting teams.live.com URLs:</h4>
          <ul className="text-sm text-gray-600 space-y-2">
            <li>
              <span className="font-medium">Use Personal Account:</span> Sign in with @outlook.com, @hotmail.com, or @live.com account
            </li>
            <li>
              <span className="font-medium">Try "Personal Teams Meeting":</span> This method specifically targets the teams.live.com format
            </li>
            <li>
              <span className="font-medium">Check Account Type:</span> Business accounts typically generate teams.microsoft.com URLs
            </li>
            <li>
              <span className="font-medium">Both Work the Same:</span> teams.live.com and teams.microsoft.com provide identical functionality
            </li>
          </ul>
          
          <h4 className="font-medium text-gray-800 mb-2 mt-4">Timezone Notes:</h4>
          <ul className="text-sm text-gray-600 space-y-2">
            <li>
              <span className="font-medium">UTC is Recommended:</span> Use UTC for international meetings to avoid confusion
            </li>
            <li>
              <span className="font-medium">Local Times Shown:</span> Teams will display meeting times in each attendee's local timezone
            </li>
            <li>
              <span className="font-medium">Daylight Saving:</span> Be aware of DST transitions when scheduling future meetings
            </li>
          </ul>
        </div>

        {/* Timezone Reference Card */}
        <div className="mt-4 p-4 bg-indigo-50 rounded-lg border border-indigo-200">
          <h4 className="font-medium text-indigo-800 mb-2">Quick Timezone Reference:</h4>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-sm">
            <div className="text-indigo-700">
              <span className="font-medium">UTC:</span> Universal time
            </div>
            <div className="text-indigo-700">
              <span className="font-medium">EST/EDT:</span> New York (-5/-4)
            </div>
            <div className="text-indigo-700">
              <span className="font-medium">PST/PDT:</span> Los Angeles (-8/-7)
            </div>
            <div className="text-indigo-700">
              <span className="font-medium">GMT/BST:</span> London (+0/+1)
            </div>
            <div className="text-indigo-700">
              <span className="font-medium">CET/CEST:</span> Paris (+1/+2)
            </div>
            <div className="text-indigo-700">
              <span className="font-medium">JST:</span> Tokyo (+9)
            </div>
          </div>
          <p className="text-xs text-indigo-600 mt-2">
            Times shown as (Standard/Daylight) offset from UTC. Your selected timezone: <strong>{formatTimezoneDisplay(timezone)}</strong>
          </p>
        </div>
      </div>
    </AuthenticatedTemplate>
  );
}
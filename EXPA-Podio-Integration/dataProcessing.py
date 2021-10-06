from graphqlclient import GraphQLClient
import json

def extract_data(lcID,startdate,enddate):
    query_ogx_APDs = """
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {
        allOpportunityApplication(
    		filters:
    		{
          person_home_lc: $lcID
          date_approved:{from:$startdate, to:$enddate}

    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      }
    }"""
    query_ogx_REs = '''
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {allOpportunityApplication(
    		filters:
    		{
          person_home_lc:$lcID
          date_realized:{from:$startdate, to:$enddate}

    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      }
    } '''
    query_ogx_FIs = '''
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {allOpportunityApplication(
    		filters:
    		{
          person_home_lc:$lcID
          experience_end_date:{from:$startdate, to:$enddate}
     statuses:["finished","completed"]
    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      } }'''

    query_icx_APDs = '''
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {allOpportunityApplication(
    		filters:
    		{
          opportunity_home_lc:$lcID
          date_approved:{from:$startdate, to:$enddate}

    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      }
    }'''
    query_icx_REs = '''
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {allOpportunityApplication(
    		filters:
    		{
          opportunity_home_lc:$lcID
          date_realized:{from:$startdate, to:$enddate}

    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      }
    } '''
    query_icx_FIs = '''
    query ($lcID: [Int],$startdate: DateTime!,$enddate: DateTime!)
    {allOpportunityApplication(
    		filters:
    		{
          opportunity_home_lc:$lcID
          experience_end_date:{from:$startdate, to:$enddate}
     statuses:["finished","completed"]
    		}

        page:1
        per_page:3000
    	)
    	{
        paging
        {
          total_items
        }
    		data
        {
          person
          {
            id
            full_name
            email

            home_lc
            {
              name
            }
            home_mc
            {
              name
            }
          }
          opportunity
          {
            id
            programme
            { short_name_display }
            host_lc
            {
              name
            }
            home_mc
            {
              name
            }


            opportunity_duration_type{
              duration_type
            }

          }
          slot{
            start_date
            end_date
          }
          status
          updated_at
          date_approved
          date_realized
          experience_end_date


         }
      } }'''
    variables = {
        "lcID": lcID,
        "startdate": startdate,
        "enddate":enddate

    }
    data = run_query(query_ogx_APDs,variables)
    result_OGX_APDs = process_data(data['data']['allOpportunityApplication']['data'], "O")

    data = run_query(query_ogx_REs,variables)
    result_OGX_REs = process_data(data['data']['allOpportunityApplication']['data'], "O")

    data = run_query(query_ogx_FIs, variables)
    result_OGX_FIs = process_data(data['data']['allOpportunityApplication']['data'], "O")

    data = run_query(query_icx_APDs, variables)
    result_ICX_APDs = process_data(data['data']['allOpportunityApplication']['data'], "I")

    data = run_query(query_icx_REs, variables)
    result_ICX_REs = process_data(data['data']['allOpportunityApplication']['data'], "I")

    data = run_query(query_icx_FIs, variables)
    result_ICX_FIs = process_data(data['data']['allOpportunityApplication']['data'], "I")

    return result_OGX_APDs, result_OGX_REs , result_OGX_FIs, result_ICX_APDs, result_ICX_REs, result_ICX_FIs


def run_query(query,variables):
    client = GraphQLClient('https://gis-api.aiesec.org/graphql?access_token=fa9a3328b00cef1125f888c1ab950903fd90aed1789366fc13182ca462630c4d')
    result = client.execute(query, variables)
    return json.loads(result)


def process_data(data,type):
    eps = []
    for ep in data:
        oneEP = []
        oneEP.append(str(ep['person']['id'])+"_"+str(ep['opportunity']['id'])) #reference ID
        oneEP.append(ep['person']['full_name'])
        oneEP.append("https://expa.aiesec.org/people/"+str(ep['person']['id']))
        oneEP.append(ep['person']['email'])
        oneEP.append(type+ep['opportunity']['programme']['short_name_display'])
        oneEP.append(ep['person']['home_lc']['name'].replace(",", " "))
        oneEP.append(ep['person']['home_mc']['name'].replace(",", " "))
        oneEP.append(ep['opportunity']['host_lc']['name'].replace(",", " "))
        oneEP.append(ep['opportunity']['home_mc']['name'].replace(",", " "))
        oneEP.append("https://expa.aiesec.org/opportunities/" + str(ep['opportunity']['id']) + "/")
        if ep['date_approved'] is None:
            oneEP.append(ep['updated_at'][:10])
        else:
            oneEP.append(ep['date_approved'][:10])
        if ep['slot'] is not None:
            oneEP.append(ep['slot']['start_date'])
            oneEP.append(ep['slot']['end_date'])
        else:
            oneEP.append("")
            oneEP.append("")
        if ep['date_realized'] is None:
            oneEP.append("")
        else:
            oneEP.append(ep['date_realized'][:10])
        if ep['experience_end_date'] is None:
            oneEP.append("")
        else:
            oneEP.append(ep['experience_end_date'][:10])

        eps.append(oneEP)

    return eps






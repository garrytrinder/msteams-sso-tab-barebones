import { Button, Title1, makeStyles, shorthands } from "@fluentui/react-components";
import { useData, useGraph } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../context";
import { useContext } from "react";

const useStyles = makeStyles({
  main: {
    ...shorthands.gap('36px'),
    ...shorthands.padding('24px'),
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'wrap'
  },
  section: {
    width: 'fit-content'
  },
});

const Tab = () => {
  const { teamsfx } = useContext(TeamsFxContext);
  const classes = useStyles();

  const profile = useGraph(
    async (graph) => {
      const profile = await graph.api("/me").get();
      return { profile };
    },
    { scope: ["User.Read"], teamsfx }
  );

  const userInfo = useData(async () => {
    if (teamsfx) {
      return await teamsfx.getUserInfo();
    }
  });

  return (
    <div className={classes.main}>
      <section className={classes.section}>
        <Title1>Welcome, {userInfo.data?.displayName}!</Title1>
      </section>
      <section className={classes.section}>
        <Button appearance="primary"
          disabled={profile.loading || profile.data?.profile}
          onClick={profile.reload}>
          Consent
        </Button>
      </section>
      {
        profile.data && (
          <section className={classes.section}>
            <pre>{JSON.stringify(profile.data, null, 2)}</pre>
          </section>
        )
      }
    </div>
  )
}

export default Tab;
